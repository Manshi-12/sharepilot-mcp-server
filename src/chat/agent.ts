import axios from "axios";
import { TOOL_SCHEMAS, executeTool } from "../tools/registry.js";

const AZURE_OPENAI_ENDPOINT = (process.env.AZURE_OPENAI_ENDPOINT || "").replace(/\/+$/, "");
const AZURE_OPENAI_API_KEY = process.env.AZURE_OPENAI_API_KEY || "";
const AZURE_OPENAI_DEPLOYMENT = process.env.AZURE_OPENAI_DEPLOYMENT || "";
const AZURE_OPENAI_API_VERSION = process.env.AZURE_OPENAI_API_VERSION || "2024-08-01-preview";

const MAX_TOOL_ROUNDS = 6; // hard cap so a confused model can't loop forever

export interface ChatMessage {
  role: "user" | "assistant" | "system";
  content: string;
}

// Convert this server's MCP-style tool schemas into the OpenAI function-calling shape.
const OPENAI_TOOLS = TOOL_SCHEMAS.map((t) => ({
  type: "function",
  function: {
    name: t.name,
    description: t.description,
    parameters: t.inputSchema,
  },
}));

const SYSTEM_PROMPT = `You are SharePilot, an AI assistant embedded in a SharePoint workspace.
You help the user search files, read documents, manage SharePoint Lists (create/read/update/delete items),
upload files and images, and create new lists — using the tools available to you.

Guidelines:
- Always use a tool when the user's request requires looking up or changing real SharePoint data. Never invent file names, list items, or IDs.
- If a tool call fails, explain what went wrong in plain language and suggest a next step — don't expose raw stack traces.
- Keep responses concise and conversational. Use bullet points or short lists when returning multiple results.
- If the request is ambiguous (e.g. which list, which library), ask a clarifying question instead of guessing.`;

function validateConfig() {
  if (!AZURE_OPENAI_ENDPOINT || !AZURE_OPENAI_API_KEY || !AZURE_OPENAI_DEPLOYMENT) {
    throw new Error(
      "Azure OpenAI is not configured. Set AZURE_OPENAI_ENDPOINT, AZURE_OPENAI_API_KEY, " +
      "and AZURE_OPENAI_DEPLOYMENT in your environment variables."
    );
  }
}

async function callAzureOpenAI(messages: any[]) {
  const url = `${AZURE_OPENAI_ENDPOINT}/openai/deployments/${AZURE_OPENAI_DEPLOYMENT}/chat/completions?api-version=${AZURE_OPENAI_API_VERSION}`;
  const { data } = await axios.post(
    url,
    {
      messages,
      tools: OPENAI_TOOLS,
      tool_choice: "auto",
      temperature: 0.3,
    },
    {
      headers: {
        "Content-Type": "application/json",
        "api-key": AZURE_OPENAI_API_KEY,
      },
      timeout: 60_000,
    }
  );
  return { message: data.choices[0].message, usage: data.usage };
}

// Streaming variant — same endpoint, `stream: true`, parsed as SSE chunks from
// Azure OpenAI itself. Used only for the final round once the model has no
// more tool calls to make, so the user sees the answer appear token-by-token.
async function streamAzureOpenAI(
  messages: any[],
  onToken: (delta: string) => void
): Promise<{ content: string; usage: any }> {
  const url = `${AZURE_OPENAI_ENDPOINT}/openai/deployments/${AZURE_OPENAI_DEPLOYMENT}/chat/completions?api-version=${AZURE_OPENAI_API_VERSION}`;
  const response = await axios.post(
    url,
    {
      messages,
      tools: OPENAI_TOOLS,
      tool_choice: "auto",
      temperature: 0.3,
      stream: true,
      stream_options: { include_usage: true },
    },
    {
      headers: {
        "Content-Type": "application/json",
        "api-key": AZURE_OPENAI_API_KEY,
      },
      timeout: 60_000,
      responseType: "stream",
    }
  );

  let content = "";
  let usage: any = null;
  let buffer = "";

  return new Promise((resolve, reject) => {
    response.data.on("data", (chunk: Buffer) => {
      buffer += chunk.toString("utf8");
      const lines = buffer.split("\n");
      buffer = lines.pop() || "";

      for (const line of lines) {
        const trimmed = line.trim();
        if (!trimmed.startsWith("data:")) continue;
        const payload = trimmed.slice(5).trim();
        if (payload === "[DONE]") continue;

        try {
          const json = JSON.parse(payload);
          const delta = json.choices?.[0]?.delta?.content;
          if (delta) {
            content += delta;
            onToken(delta);
          }
          if (json.usage) usage = json.usage;
        } catch {
          // Partial/malformed chunk boundary — ignore, next chunk completes it
        }
      }
    });

    response.data.on("end", () => resolve({ content, usage }));
    response.data.on("error", (err: Error) => reject(err));
  });
}

// ── Event shapes sent back over SSE to the caller (the /chat route) ─────────
// "token"      — one streamed text chunk of the final answer
// "tool_call"  — the model decided to call a tool, with its parsed arguments
// "tool_result"— that tool's result (or error), paired to the same call id
// "usage"      — token usage for one round-trip to Azure OpenAI
// "done"       — final event, carries the complete assistant message
export type AgentEvent =
  | { type: "tool_call"; id: string; name: string; arguments: any; round: number }
  | { type: "tool_result"; id: string; name: string; result: any; isError: boolean; round: number }
  | { type: "token"; delta: string }
  | { type: "usage"; promptTokens: number; completionTokens: number; totalTokens: number; round: number }
  | { type: "done"; content: string };

/**
 * Runs the full agent loop: sends the conversation to Azure OpenAI, executes
 * any tool calls the model requests against the real SharePoint tools, feeds
 * the results back, and repeats until the model returns a plain-text answer.
 * Emits every step as a real event via `onEvent` so callers can stream
 * progress (tool calls, tool results, token usage, streamed final tokens)
 * to the frontend instead of waiting for one big response at the end.
 */
export async function runChatAgent(
  history: ChatMessage[],
  onEvent?: (event: AgentEvent) => void
): Promise<string> {
  validateConfig();
  const emit = onEvent || (() => { });

  const messages: any[] = [
    { role: "system", content: SYSTEM_PROMPT },
    ...history.map((m) => ({ role: m.role, content: m.content })),
  ];

  for (let round = 0; round < MAX_TOOL_ROUNDS; round++) {
    const isFinalAttempt = round === MAX_TOOL_ROUNDS - 1;

    const { message: assistantMessage, usage } = await callAzureOpenAI(messages);

    if (usage) {
      emit({
        type: "usage",
        promptTokens: usage.prompt_tokens ?? 0,
        completionTokens: usage.completion_tokens ?? 0,
        totalTokens: usage.total_tokens ?? 0,
        round,
      });
    }

    const toolCalls = assistantMessage.tool_calls;

    if (!toolCalls || toolCalls.length === 0) {
      // Final answer — re-run as a stream so tokens arrive live. Same
      // `messages` array in, so the model deterministically reproduces the
      // same answer, just delivered incrementally this time.
      const streamed = await streamAzureOpenAI(messages, (delta) => emit({ type: "token", delta }));
      const finalContent = streamed.content || assistantMessage.content || "I couldn't generate a response. Please try again.";
      emit({ type: "done", content: finalContent });
      return finalContent;
    }

    messages.push(assistantMessage);

    for (const call of toolCalls) {
      let args: any = {};
      try {
        args = call.function.arguments ? JSON.parse(call.function.arguments) : {};
      } catch {
        // Model produced malformed JSON args — surface that back to it
      }

      emit({ type: "tool_call", id: call.id, name: call.function.name, arguments: args, round });

      let resultText: string;
      let resultForEvent: any;
      let isError = false;
      try {
        const result = await executeTool(call.function.name, args);
        resultText = JSON.stringify(result, null, 2);
        resultForEvent = result;
      } catch (err: any) {
        const graphError = err?.response?.data ? JSON.stringify(err.response.data) : err.message || String(err);
        resultText = `Error executing tool "${call.function.name}": ${graphError}`;
        resultForEvent = { error: graphError };
        isError = true;
      }

      emit({ type: "tool_result", id: call.id, name: call.function.name, result: resultForEvent, isError, round });

      messages.push({
        role: "tool",
        tool_call_id: call.id,
        content: resultText,
      });
    }

    if (isFinalAttempt) {
      const fallback = "I wasn't able to finish that request after several tool calls — please try rephrasing or breaking it into smaller steps.";
      emit({ type: "done", content: fallback });
      return fallback;
    }
  }

  const fallback = "I wasn't able to finish that request after several tool calls — please try rephrasing or breaking it into smaller steps.";
  emit({ type: "done", content: fallback });
  return fallback;
}
