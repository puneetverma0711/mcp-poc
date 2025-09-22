import {
  McpServer,
  ResourceTemplate,
} from "@modelcontextprotocol/sdk/server/mcp.js"
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js"
import { z } from "zod"
import fs from "node:fs/promises"
import path from "node:path"
import { pathToFileURL } from "node:url"
import { CreateMessageResultSchema } from "@modelcontextprotocol/sdk/types.js"
import * as cheerio from "cheerio"; // for HTML parsing


 

const server = new McpServer({
  name: "azure-mcp",
  version: "1.0.0",
  capabilities: {
    resources: {},
    tools: {},
    prompts: {},
  },
})


server.prompt(
  "get-issues-by-state",
  "Gets a list of Azure DevOps issues by their state.",
  {
    state: z.string().describe("The state of the issues to retrieve (e.g., 'To Do', 'Doing', 'Done')."),
  },
  async ({ state }) => {
    try {
      const data = await fs.readFile("src/data/azzureissues.json", "utf-8");
      const issues = JSON.parse(data);
      const filteredIssues = issues.filter(
        (issue: any) => issue.fields["System.State"] === state
      );
      const issueTitles = filteredIssues.map((issue: any) => `- ${issue.fields["System.Title"]}`).join("\n");
      return {
        messages: [
          {
            role: "assistant",
            content: {
              type: "text",
              text: `Issues in state "${state}":\n${issueTitles}`,
            },
          },
        ],
      };
    } catch (error) {
      return {
        messages: [
          {
            role: "assistant",
            content: {
              type: "text",
              text: `Operation failed: ${error}`,
            },
          },
        ],
      };
    }
  }
);

server.tool(
  "getazzureIssuesid",
  "Fetches Azure DevOps issue IDs, gets their details, saves them to a file, and generates an HTML report.",
  {
    username: z.string().describe("Azure DevOps username (often email)"),
    pat: z.string().describe("Azure DevOps Personal Access Token"),
  },
  {
    title: "Get Azure Issues and Report",
    readOnlyHint: false,
    destructiveHint: false,
    idempotentHint: false,
    openWorldHint: true,
  },
  async ({ username, pat }) => {
    const authHeader = `Basic ${Buffer.from(`${username}:${pat}`).toString(
      "base64"
    )}`
    const headers = {
      "Content-Type": "application/json",
      Authorization: authHeader,
    }

    try {
      // Step 1: Get Issue IDs
      const wiqlUrl = `https://dev.azure.com/CollabMCP/Collabro/_apis/wit/wiql?api-version=7.1`
      const wiqlQuery = {
        query:
          "SELECT [System.Id] FROM WorkItems WHERE [System.TeamProject] = 'Collabro' AND [System.WorkItemType] = 'Issue' ORDER BY [System.ChangedDate] DESC",
      }
      const wiqlResponse = await fetch(wiqlUrl, {
        method: "POST",
        headers,
        body: JSON.stringify(wiqlQuery),
      })

      if (!wiqlResponse.ok) {
        throw new Error(
          `Failed to fetch issue IDs: ${
            wiqlResponse.status
          } ${await wiqlResponse.text()}`
        )
      }



      const wiqlResult = await wiqlResponse.json()
      const issueIds = wiqlResult.workItems.map((item: any) => item.id)

      if (issueIds.length === 0) {
        return { content: [{ type: "text", text: "No issues found." }] }
      }

      // Step 2: Get Issue Details
      const workItemsUrl = `https://dev.azure.com/CollabMCP/Collabro/_apis/wit/workitemsbatch?api-version=7.1`
      const workItemsPayload = {
        ids: issueIds,
        fields: [
          "System.Id",
          "System.Title",
          "System.State",
          "System.AssignedTo",
          "System.CreatedDate",
          "System.ChangedDate",
          "System.Tags",
        ],
      }
      const workItemsResponse = await fetch(workItemsUrl, {
        method: "POST",
        headers,
        body: JSON.stringify(workItemsPayload),
      })

      if (!workItemsResponse.ok) {
        throw new Error(
          `Failed to fetch issue details: ${
            workItemsResponse.status
          } ${await workItemsResponse.text()}`
        )
      }

      const workItemsDetails = await workItemsResponse.json();
      const jsonFilePath = "src/data/azzureissues.json"
      await fs.writeFile( jsonFilePath,JSON.stringify(workItemsDetails.value, null, 2)
      )

      return {
        content: [
          {
            type: "text",
            text: `Successfully fetched ${workItemsDetails.count} issues. Data saved to ${jsonFilePath}.`,
          },
        ],
      }
    } catch (error) {
      return {
        content: [
          {
            type: "text",
            text: `Operation failed: ${error}`,
          },
        ],
      }
    }
  }
)

server.tool(
  "getIssuesDetailsById",
  "Gets details for a specific list of Azure DevOps work item IDs.",
  {
    username: z.string().describe("Azure DevOps username (often email)"),
    pat: z.string().describe("Azure DevOps Personal Access Token"),
    ids: z
      .array(z.number())
      .describe("Array of work item IDs to fetch details for."),
  },
  {
    title: "Get Azure Issue Details by ID",
    readOnlyHint: true,
    destructiveHint: false,
    idempotentHint: true,
    openWorldHint: true,
  },
  async ({ username, pat, ids }) => {
    const authHeader = `Basic ${Buffer.from(`${username}:${pat}`).toString(
      "base64"
    )}`
    const headers = {
      "Content-Type": "application/json",
      Authorization: authHeader,
    }

    try {
      const workItemsUrl = `https://dev.azure.com/CollabMCP/Collabro/_apis/wit/workitemsbatch?api-version=7.1`
      const workItemsPayload = {
        ids: ids,
        fields: [
          "System.Id",
          "System.Title",
          "System.State",
          "System.AssignedTo",
          "System.CreatedDate",
          "System.ChangedDate",
          "System.Tags",
        ],
      }
      const workItemsResponse = await fetch(workItemsUrl, {
        method: "POST",
        headers,
        body: JSON.stringify(workItemsPayload),
      })

      if (!workItemsResponse.ok) {
        throw new Error(
          `Failed to fetch issue details: ${
            workItemsResponse.status
          } ${await workItemsResponse.text()}`
        )
      }

      const workItemsDetails = await workItemsResponse.json()
      return {
        content: [
          {
            type: "text",
            text: JSON.stringify(workItemsDetails, null, 2),
          },
        ],
      }
    } catch (error) {
      return {
        content: [
          {
            type: "text",
            text: `Operation failed: ${error}`,
          },
        ],
      }
    }
  }
)


server.tool(
  "api-request",
  "Make a generic API request to any endpoint",
  {
    url: z.string().url(),
    method: z.enum(["GET", "POST", "PUT", "DELETE", "PATCH"]),
    headers: z.record(z.string()).optional(),
    body: z.any().optional(),
  },
  {
    title: "API Request",
    readOnlyHint: false,
    destructiveHint: false,
    idempotentHint: false,
    openWorldHint: true,
  },
  async ({ url, method, headers, body }) => {
    try {
      const response = await fetch(url, {
        method,
        headers,
        body: body ? JSON.stringify(body) : undefined,
      })
      const contentType = response.headers.get("content-type") || ""
      let data
      if (contentType.includes("application/json")) {
        data = await response.json()
      } else {
        data = await response.text()
      }
      return {
        content: [
          {
            type: "text",
            text: `Status: ${response.status}\nResponse: ${JSON.stringify(data, null, 2)}`,
          },
        ],
      }
    } catch (error) {
      return {
        content: [
          {
            type: "text",
            text: `API request failed: ${error}`,
          },
        ],
      }
    }
  }
)



async function main() {
  const transport = new StdioServerTransport()
  await server.connect(transport)
  console.log("Server started");
}

main()

