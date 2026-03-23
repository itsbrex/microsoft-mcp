/**
 * inbox_triage.ts
 *
 * Code Mode example: fetch ranked inbox summaries, hydrate the top 3 items,
 * and return a compact triage report.
 *
 * This script demonstrates the recommended orchestration pattern for inbox
 * triage using the microsoft-mcp server:
 *
 *   1. list_inbox_items  — low-cost ranked summaries
 *   2. get_inbox_item_detail — hydrate only the selected items
 *   3. Compute and return a triage report locally in Code Mode
 *
 * Code Mode is used here to batch the follow-up detail calls and reduce
 * them to a single compact output. The server handles response shaping
 * (trimming raw Graph payloads). Code Mode handles orchestration logic.
 *
 * Prerequisites:
 *   - microsoft-mcp registered in your MCP config (see docs/code-mode-inbox-orchestration.md)
 *   - Authentication completed for the configured account
 */

import Anthropic from "@anthropic-ai/sdk";

// ---------------------------------------------------------------------------
// MCP server configuration
// Update paths and account details to match your local setup.
// ---------------------------------------------------------------------------
const MCP_SERVER_CONFIG = {
  mcpServers: {
    "microsoft-mcp": {
      command: process.env.UV_PATH ?? "uv",
      args: [
        "run",
        "--python",
        "3.13",
        "--project",
        process.env.MICROSOFT_MCP_PROJECT_PATH ?? "/path/to/microsoft-mcp",
        "microsoft-mcp",
      ],
      env: {
        MICROSOFT_MCP_AUTH_METHOD: "msal",
        MICROSOFT_MCP_ACCOUNT_ID:
          process.env.MICROSOFT_MCP_ACCOUNT_ID ?? "your-email@example.com",
        MICROSOFT_MCP_CLIENT_ID:
          process.env.MICROSOFT_MCP_CLIENT_ID ??
          "d3590ed6-52b3-4102-aeff-aad2292ab01c",
      },
    },
  },
};

// ---------------------------------------------------------------------------
// Types mirroring InboxItem fields returned by list_inbox_items
// ---------------------------------------------------------------------------
interface InboxItem {
  id: string;
  kind: "email" | "event";
  source_tool: string;
  title: string;
  snippet: string;
  participants: string[];
  when: string;
  state: string;
  score: number;
  reason: string;
  action_hints: string[];
  web_url: string;
}

interface InboxListResult {
  items: InboxItem[];
  total: number;
}

interface ItemDetail {
  id: string;
  kind: string;
  title: string;
  body: string;
  participants: string[];
  when: string;
  action_hints: string[];
  web_url: string;
}

// ---------------------------------------------------------------------------
// Triage report entry
// ---------------------------------------------------------------------------
interface TriageEntry {
  title: string;
  kind: string;
  summary: string;
  suggested_action: string;
  urgency_reason: string;
  web_url: string;
}

// ---------------------------------------------------------------------------
// Helper: call an MCP tool and parse the JSON result
// ---------------------------------------------------------------------------
async function callMcpTool<T>(
  client: Anthropic,
  toolName: string,
  toolInput: Record<string, unknown>
): Promise<T> {
  // In a real Code Mode script this would be a direct MCP tool invocation.
  // Here we use the Anthropic client with tool_choice to demonstrate the pattern.
  const response = await client.messages.create({
    model: "claude-sonnet-4-5",
    max_tokens: 4096,
    tools: [
      {
        name: toolName,
        description: `Call ${toolName}`,
        input_schema: {
          type: "object" as const,
          properties: toolInput,
        },
      },
    ],
    tool_choice: { type: "any" },
    messages: [
      {
        role: "user",
        content: `Call ${toolName} with the provided parameters: ${JSON.stringify(toolInput)}`,
      },
    ],
  });

  // Extract tool result from response
  for (const block of response.content) {
    if (block.type === "tool_use") {
      return block.input as T;
    }
  }
  throw new Error(`No tool_use block in response from ${toolName}`);
}

// ---------------------------------------------------------------------------
// Build a one-sentence summary from item detail
// ---------------------------------------------------------------------------
function buildSummary(detail: ItemDetail): string {
  const bodyPreview = detail.body
    ? detail.body.slice(0, 150).replace(/\s+/g, " ").trim()
    : "(no body)";
  return bodyPreview.length < detail.body?.length
    ? bodyPreview + "..."
    : bodyPreview;
}

// ---------------------------------------------------------------------------
// Suggest an action based on item kind and action_hints
// ---------------------------------------------------------------------------
function suggestAction(detail: ItemDetail): string {
  if (detail.action_hints && detail.action_hints.length > 0) {
    return detail.action_hints[0];
  }
  if (detail.kind === "event") {
    return "Review and accept or decline";
  }
  return "Read and reply if needed";
}

// ---------------------------------------------------------------------------
// Main triage orchestration
// ---------------------------------------------------------------------------
async function runInboxTriage(): Promise<void> {
  console.log("Starting inbox triage...\n");

  // In a Code Mode script registered against the MCP server, tool calls happen
  // via the MCP protocol. This example shows the logical flow using the
  // Anthropic SDK as a stand-in for illustration.
  //
  // The key pattern is:
  //   1. list_inbox_items -> ranked summaries (one call, low cost)
  //   2. select top N by score
  //   3. get_inbox_item_detail for each selected item (N calls, only for top items)
  //   4. build report locally — no additional model calls needed

  // Step 1: Fetch ranked inbox summaries
  console.log("Step 1: Fetching ranked inbox summaries (limit=20)...");
  // In production Code Mode: const summaryResult = await mcp.list_inbox_items({ limit: 20 });
  // Simulated result for demonstration:
  const summaryResult: InboxListResult = {
    items: [
      {
        id: "item-001",
        kind: "email",
        source_tool: "list_emails",
        title: "Q2 budget approval needed by EOD",
        snippet: "Hi, please review and approve the attached budget before 5pm today.",
        participants: ["cfo@company.com", "you@company.com"],
        when: "2026-03-23T08:15:00Z",
        state: "unread",
        score: 0.92,
        reason: "Unread, flagged, deadline today",
        action_hints: ["Reply with approval", "Forward to finance team"],
        web_url: "https://outlook.office.com/mail/id/item-001",
      },
      {
        id: "item-002",
        kind: "event",
        source_tool: "list_events",
        title: "All-hands meeting in 30 minutes",
        snippet: "Quarterly all-hands. Please confirm attendance.",
        participants: ["organizer@company.com"],
        when: "2026-03-23T10:00:00Z",
        state: "tentative",
        score: 0.88,
        reason: "Meeting in <1 hour, response pending",
        action_hints: ["Accept or decline"],
        web_url: "https://outlook.office.com/calendar/id/item-002",
      },
      {
        id: "item-003",
        kind: "email",
        source_tool: "list_emails",
        title: "You were mentioned in #incident-channel",
        snippet: "Hey @you, can you check the deployment logs for the auth service?",
        participants: ["devlead@company.com"],
        when: "2026-03-23T09:45:00Z",
        state: "unread",
        score: 0.81,
        reason: "Unread, direct mention",
        action_hints: ["Check logs and reply"],
        web_url: "https://outlook.office.com/mail/id/item-003",
      },
    ],
    total: 3,
  };

  console.log(`  Received ${summaryResult.total} items. Top scores:`);
  summaryResult.items.forEach((item, i) => {
    console.log(
      `  ${i + 1}. [${item.kind}] "${item.title}" — score=${item.score.toFixed(2)} (${item.reason})`
    );
  });

  // Step 2: Select top 3 items for hydration (already sorted by score)
  const TOP_N = 3;
  const selectedItems = summaryResult.items.slice(0, TOP_N);
  console.log(`\nStep 2: Hydrating top ${TOP_N} items...`);

  // Step 3: Hydrate selected items (parallel calls)
  // In production Code Mode:
  //   const details = await Promise.all(
  //     selectedItems.map(item =>
  //       mcp.get_inbox_item_detail({ item_id: item.id, kind: item.kind })
  //     )
  //   );
  // Simulated details for demonstration:
  const details: ItemDetail[] = selectedItems.map((item) => ({
    id: item.id,
    kind: item.kind,
    title: item.title,
    body:
      item.kind === "email"
        ? `Full email body for "${item.title}". ${item.snippet} [Additional context from full body...]`
        : `Event description: ${item.snippet}`,
    participants: item.participants,
    when: item.when,
    action_hints: item.action_hints,
    web_url: item.web_url,
  }));

  console.log(`  Hydrated ${details.length} items.`);

  // Step 4: Build triage report locally in Code Mode
  console.log("\nStep 4: Building triage report...\n");
  const report: TriageEntry[] = details.map((detail, i) => ({
    title: detail.title,
    kind: detail.kind,
    summary: buildSummary(detail),
    suggested_action: suggestAction(detail),
    urgency_reason: selectedItems[i].reason,
    web_url: detail.web_url,
  }));

  // Output the compact triage report
  console.log("=".repeat(60));
  console.log("INBOX TRIAGE REPORT");
  console.log("=".repeat(60));
  report.forEach((entry, i) => {
    console.log(`\n${i + 1}. [${entry.kind.toUpperCase()}] ${entry.title}`);
    console.log(`   Why urgent: ${entry.urgency_reason}`);
    console.log(`   Summary: ${entry.summary}`);
    console.log(`   Suggested action: ${entry.suggested_action}`);
    console.log(`   Link: ${entry.web_url}`);
  });
  console.log("\n" + "=".repeat(60));
  console.log(`Triage complete. ${report.length} items require attention.`);
}

// ---------------------------------------------------------------------------
// Entry point
// ---------------------------------------------------------------------------
runInboxTriage().catch((err) => {
  console.error("Triage failed:", err);
  process.exit(1);
});
