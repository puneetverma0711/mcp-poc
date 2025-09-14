import {
  McpServer,
  ResourceTemplate,
} from "@modelcontextprotocol/sdk/server/mcp.js"
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js"
import { z } from "zod"
import fs from "node:fs/promises"
import { CreateMessageResultSchema } from "@modelcontextprotocol/sdk/types.js"

const server = new McpServer({
  name: "azure-mcp",
  version: "1.0.0",
  capabilities: {
    resources: {},
    tools: {},
    prompts: {},
  },
})



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

server.tool(
  "azure-api-get",
  "Make a GET request to an Azure DevOps API endpoint using Basic Auth",
  {
    url: z.string().url(),
    username: z.string(),
    password: z.string(),
    headers: z.record(z.string()).optional(),
  },
  {
    title: "Azure API GET Request",
    readOnlyHint: true,
    destructiveHint: false,
    idempotentHint: true,
    openWorldHint: false,
  },
  async ({ url, username, password, headers }) => {
    try {
      const basicAuth = Buffer.from(`${username}:${password}`).toString("base64")
      const response = await fetch(url, {
        method: "GET",
        headers: {
          ...(headers || {}),
          Authorization: `Basic ${basicAuth}`,
        },
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
            text: `Azure API request failed: ${error}`,
          },
        ],
      }
    }
  }
)

server.tool(
  "azure-api-get-and-post-teams",
  "Fetch data from Azure DevOps API and post it to a Microsoft Teams channel",
  {
    url: z.string().url(),
    username: z.string().optional(),
    password: z.string().optional(),
    teamsWebhookUrl: z.string().url(),
    headers: z.record(z.string()).optional(),
  },
  {
    title: "Azure API GET and Post to Teams",
    readOnlyHint: false,
    destructiveHint: false,
    idempotentHint: false,
    openWorldHint: false,
  },
  async ({ url, username, password, teamsWebhookUrl, headers }) => {
    try {
      // Use env vars if username/password not provided
      const user = username || process.env.AZURE_USERNAME || ""
      const pass = password || process.env.AZURE_PASSWORD || ""
      const basicAuth = Buffer.from(`${user}:${pass}`).toString("base64")
      // Fetch from Azure DevOps
      const response = await fetch(url, {
        method: "GET",
        headers: {
          ...(headers || {}),
          Authorization: `Basic ${basicAuth}`,
        },
      })
      const contentType = response.headers.get("content-type") || ""
      let data
      if (contentType.includes("application/json")) {
        data = await response.json()
      } else {
        data = await response.text()
      }
      // Format message for Teams
      const teamsMessage = {
        text: `Azure DevOps API Response:\n\`\`\`json\n${JSON.stringify(data, null, 2)}\n\`\`\``
      }
      // Post to Teams
      const teamsRes = await fetch(teamsWebhookUrl, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(teamsMessage),
      })
      if (!teamsRes.ok) {
        return {
          content: [
            {
              type: "text",
              text: `Failed to post to Teams: ${teamsRes.status} ${teamsRes.statusText}`,
            },
          ],
        }
      }
      return {
        content: [
          {
            type: "text",
            text: `Successfully fetched from Azure DevOps and posted to Teams.`,
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




server.prompt(
  "get-azure-workitem",
  "Prompt to get an Azure DevOps work item by ID",
  {
    id: z.string(),
  },
  ({ id }) => ({
    messages: [
      {
        role: "user",
        content: {
          type: "text",
          text: `Fetch the Azure DevOps work item with ID ${id}.`,
        },
      },
    ],
  })
)


server.resource(
  "azure-workitem",
  new ResourceTemplate("azure://workitems/{id}", { list: undefined }),
  {
    description: "Get an Azure DevOps work item by ID",
    title: "Azure Work Item",
    mimeType: "application/json",
  },
  async (uri, { id }) => {
    // You may want to store credentials securely or fetch from env/config
    const url = `https://dev.azure.com/CollabMCP/Collabro/_apis/wit/workitems/${id}?api-version=7.1`
    const username = process.env.AZURE_USERNAME || ""
    const password = process.env.AZURE_PASSWORD || ""
    const basicAuth = Buffer.from(`${username}:${password}`).toString("base64")

    try {
      const response = await fetch(url, {
        method: "GET",
        headers: {
          Authorization: `Basic ${basicAuth}`,
        },
      })
      const data = await response.json()
      return {
        contents: [
          {
            uri: uri.href,
            text: JSON.stringify(data, null, 2),
            mimeType: "application/json",
          },
        ],
      }
    } catch (error) {
      return {
        contents: [
          {
            uri: uri.href,
            text: JSON.stringify({ error: String(error) }),
            mimeType: "application/json",
          },
        ],
      }
    }
  }
)

// ...existing code...

async function main() {
  const transport = new StdioServerTransport()
  await server.connect(transport)
}

main()
