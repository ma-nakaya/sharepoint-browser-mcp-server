import assert from "node:assert/strict";
import test from "node:test";

import { Client } from "@modelcontextprotocol/sdk/client/index.js";
import { InMemoryTransport } from "@modelcontextprotocol/sdk/inMemory.js";
import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import {
  CallToolResultSchema,
  type CallToolResult,
} from "@modelcontextprotocol/sdk/types.js";

import type { KnowledgeRetrievalService } from "../src/sharepoint/knowledge-retrieval-service.js";
import { registerKnowledgeRetrievalTools } from "../src/tools/knowledge-retrieval-tools.js";

test("advertises portable search and fetch tools over MCP", async () => {
  const pageUrl =
    "https://example.sharepoint.com/sites/genai/SitePages/Home.aspx";
  const service = {
    search: async () => ({
      results: [{ id: pageUrl, title: "Home", url: pageUrl }],
    }),
    fetch: async () => ({
      id: pageUrl,
      title: "Home",
      text: "Welcome.",
      url: pageUrl,
      metadata: {
        source: "sharepoint",
        kind: "page",
        truncated: false,
      },
    }),
  };
  const server = new McpServer({
    name: "knowledge-retrieval-test",
    version: "0.0.0",
  });
  registerKnowledgeRetrievalTools(
    server,
    service as unknown as KnowledgeRetrievalService,
  );

  const client = new Client({
    name: "knowledge-retrieval-test-client",
    version: "0.0.0",
  });
  const [clientTransport, serverTransport] =
    InMemoryTransport.createLinkedPair();

  try {
    await Promise.all([
      server.connect(serverTransport),
      client.connect(clientTransport),
    ]);

    const tools = await client.listTools();
    const searchTool = tools.tools.find((tool) => tool.name === "search");
    const fetchTool = tools.tools.find((tool) => tool.name === "fetch");
    assert.ok(searchTool);
    assert.ok(fetchTool);
    assert.equal(searchTool.annotations?.readOnlyHint, true);
    assert.equal(fetchTool.annotations?.readOnlyHint, true);
    assert.deepEqual(searchTool.inputSchema.required, ["query"]);
    assert.deepEqual(fetchTool.inputSchema.required, ["id"]);

    const searchResult = await client.callTool({
      name: "search",
      arguments: { query: "home" },
    }, CallToolResultSchema);
    const validatedSearchResult = CallToolResultSchema.parse(searchResult);
    assert.deepEqual(validatedSearchResult.structuredContent, {
      results: [{ id: pageUrl, title: "Home", url: pageUrl }],
    });
    assert.deepEqual(JSON.parse(getTextContent(validatedSearchResult.content)), {
      results: [{ id: pageUrl, title: "Home", url: pageUrl }],
    });

    const fetchResult = await client.callTool({
      name: "fetch",
      arguments: { id: pageUrl },
    }, CallToolResultSchema);
    const validatedFetchResult = CallToolResultSchema.parse(fetchResult);
    assert.deepEqual(validatedFetchResult.structuredContent, {
      id: pageUrl,
      title: "Home",
      text: "Welcome.",
      url: pageUrl,
      metadata: {
        source: "sharepoint",
        kind: "page",
        truncated: false,
      },
    });
    assert.deepEqual(
      JSON.parse(getTextContent(validatedFetchResult.content)),
      validatedFetchResult.structuredContent,
    );
  } finally {
    await client.close();
    await server.close();
  }
});

test("returns safe fetch failures as MCP tool errors", async () => {
  const service = {
    search: async () => ({ results: [] }),
    fetch: async () => {
      throw new Error("The SharePoint list item was not found.");
    },
  };
  const server = new McpServer({
    name: "knowledge-retrieval-error-test",
    version: "0.0.0",
  });
  registerKnowledgeRetrievalTools(
    server,
    service as unknown as KnowledgeRetrievalService,
  );
  const client = new Client({
    name: "knowledge-retrieval-error-test-client",
    version: "0.0.0",
  });
  const [clientTransport, serverTransport] =
    InMemoryTransport.createLinkedPair();

  try {
    await Promise.all([
      server.connect(serverTransport),
      client.connect(clientTransport),
    ]);

    const result = await client.callTool({
      name: "fetch",
      arguments: { id: pageUrlForErrorTest },
    });
    const validatedResult = CallToolResultSchema.parse(result);

    assert.equal(validatedResult.isError, true);
    assert.equal(
      getTextContent(validatedResult.content),
      "The SharePoint list item was not found.",
    );
  } finally {
    await client.close();
    await server.close();
  }
});

const pageUrlForErrorTest =
  "https://example.sharepoint.com/sites/genai/Lists/Rules/DispForm.aspx?ID=1";

function getTextContent(
  content: CallToolResult["content"],
): string {
  const item = content[0];
  assert.ok(item);
  assert.equal(item.type, "text");
  if (item.type !== "text") {
    throw new Error("Expected text content.");
  }
  return item.text;
}
