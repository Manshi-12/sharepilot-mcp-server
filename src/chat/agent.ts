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
  return data.choices[0].message;
}

/**
 * Runs the full agent loop: sends the conversation to Azure OpenAI, executes
 * any tool calls the model requests against the real SharePoint tools, feeds
 * the results back, and repeats until the model returns a plain-text answer.
 */
export async function runChatAgent(history: ChatMessage[]): Promise<string> {
  validateConfig();

  const messages: any[] = [
    { role: "system", content: SYSTEM_PROMPT },
    ...history.map((m) => ({ role: m.role, content: m.content })),
  ];

  for (let round = 0; round < MAX_TOOL_ROUNDS; round++) {
    const assistantMessage = await callAzureOpenAI(messages);
    messages.push(assistantMessage);

    const toolCalls = assistantMessage.tool_calls;
    if (!toolCalls || toolCalls.length === 0) {
      return assistantMessage.content || "I couldn't generate a response. Please try again.";
    }

    for (const call of toolCalls) {
      let args: any = {};
      try {
        args = call.function.arguments ? JSON.parse(call.function.arguments) : {};
      } catch {
        // Model produced malformed JSON args — surface that back to it
      }

      let resultText: string;
      try {
        const result = await executeTool(call.function.name, args);
        resultText = JSON.stringify(result, null, 2);
      } catch (err: any) {
        const graphError = err?.response?.data ? JSON.stringify(err.response.data) : err.message || String(err);
        resultText = `Error executing tool "${call.function.name}": ${graphError}`;
      }

      messages.push({
        role: "tool",
        tool_call_id: call.id,
        content: resultText,
      });
    }
  }

  return "I wasn't able to finish that request after several tool calls — please try rephrasing or breaking it into smaller steps.";
}
