import axios from "axios";
import { TOOL_SCHEMAS, executeTool } from "../tools/registry.js";
import { SYSTEM_PROMPT } from "./systemPrompt.js";

const AZURE_OPENAI_ENDPOINT = (process.env.AZURE_OPENAI_ENDPOINT || "").replace(/\/+$/, "");
const AZURE_OPENAI_API_KEY = process.env.AZURE_OPENAI_API_KEY || "";
const AZURE_OPENAI_DEPLOYMENT = process.env.AZURE_OPENAI_DEPLOYMENT || "";
const AZURE_OPENAI_API_VERSION = process.env.AZURE_OPENAI_API_VERSION || "2024-08-01-preview";

const MAX_TOOL_ROUNDS = 6;

export interface ChatMessage {
  role: "user" | "assistant" | "system";
  content: string;
}

const OPENAI_TOOLS = TOOL_SCHEMAS.map((t) => ({
  type: "function",
  function: {
    name: t.name,
    description: t.description,
    parameters: t.inputSchema,
  },
}));

function validateConfig() {
  if (!AZURE_OPENAI_ENDPOINT || !AZURE_OPENAI_API_KEY || !AZURE_OPENAI_DEPLOYMENT) {
    throw new Error(
      "Azure OpenAI is not configured. Set AZURE_OPENAI_ENDPOINT, AZURE_OPENAI_API_KEY, " +
      "and AZURE_OPENAI_DEPLOYMENT in your environment variables."
    );
  }
}

// ── Single streaming call that handles BOTH tool calls and final text ─────────
// Streams the response from Azure OpenAI. If the model returns tool calls,
// they are accumulated from the stream and returned. If it returns text,
// each token is emitted immediately via onToken so the user sees it live.
async function streamAzureOpenAI(
  messages: any[],
  onToken: (delta: string) => void
): Promise<{ toolCalls: any[] | null; content: string; usage: any }> {
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

  // Accumulate tool call deltas — Azure streams them in fragments
  const toolCallMap: Record<number, { id: string; name: string; arguments: string }> = {};

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
          const delta = json.choices?.[0]?.delta;

          if (delta?.content) {
            content += delta.content;
            // Emit each token immediately — no buffering
            onToken(delta.content);
          }

          // Accumulate tool call fragments
          if (delta?.tool_calls) {
            for (const tc of delta.tool_calls) {
              const idx = tc.index ?? 0;
              if (!toolCallMap[idx]) {
                toolCallMap[idx] = { id: "", name: "", arguments: "" };
              }
              if (tc.id) toolCallMap[idx].id = tc.id;
              if (tc.function?.name) toolCallMap[idx].name += tc.function.name;
              if (tc.function?.arguments) toolCallMap[idx].arguments += tc.function.arguments;
            }
          }

          if (json.usage) usage = json.usage;
        } catch {
          // Partial/malformed chunk — ignore, next chunk completes it
        }
      }
    });

    response.data.on("end", () => {
      const toolCalls = Object.keys(toolCallMap).length > 0
        ? Object.values(toolCallMap).map((tc) => ({
          id: tc.id,
          function: { name: tc.name, arguments: tc.arguments },
        }))
        : null;

      resolve({ toolCalls, content, usage });
    });

    response.data.on("error", (err: Error) => reject(err));
  });
}

export type AgentEvent =
  | { type: "tool_call"; id: string; name: string; arguments: any; round: number }
  | { type: "tool_result"; id: string; name: string; result: any; isError: boolean; round: number }
  | { type: "token"; delta: string }
  | { type: "usage"; promptTokens: number; completionTokens: number; totalTokens: number; round: number }
  | { type: "done"; content: string };

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

    // Single streaming call — tokens flow immediately if it's a text answer,
    // or tool call fragments are accumulated if the model wants to use a tool.
    const { toolCalls, content, usage } = await streamAzureOpenAI(
      messages,
      (delta) => emit({ type: "token", delta })
    );

    if (usage) {
      emit({
        type: "usage",
        promptTokens: usage.prompt_tokens ?? 0,
        completionTokens: usage.completion_tokens ?? 0,
        totalTokens: usage.total_tokens ?? 0,
        round,
      });
    }

    // No tool calls → the streamed text IS the final answer
    if (!toolCalls || toolCalls.length === 0) {
      const finalContent = content || "I couldn't generate a response. Please try again.";
      emit({ type: "done", content: finalContent });
      return finalContent;
    }

    // Tool calls — push the assistant message with tool_calls, execute each tool
    messages.push({
      role: "assistant",
      content: content || null,
      tool_calls: toolCalls.map((tc) => ({
        id: tc.id,
        type: "function",
        function: tc.function,
      })),
    });

    for (const call of toolCalls) {
      let args: any = {};
      try {
        args = call.function.arguments ? JSON.parse(call.function.arguments) : {};
      } catch {
        // Malformed JSON args from model
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
        const graphError = err?.response?.data
          ? JSON.stringify(err.response.data)
          : err.message || String(err);
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