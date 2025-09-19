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

      // Step 4: Generate HTML report
      let htmlReport = `
        <html>
        <head>
          <title>Azure DevOps Work Item Report</title>
          <style>
            body { font-family: sans-serif; }
            table { border-collapse: collapse; width: 100%; }
            th, td { border: 1px solid #ddd; padding: 8px; text-align: left; }
            th { background-color: #f2f2f2; }
          </style>
        </head>
        <body>
          <h1>Azure DevOps Work Item Report</h1>
          <table>
            <tr>
              <th>ID</th>
              <th>Title</th>
              <th>State</th>
              <th>Assigned To</th>
              <th>Created Date</th>
              <th>Changed Date</th>
              <th>Tags</th>
            </tr>`

      for (const item of workItemsDetails.value) {
        htmlReport += `
            <tr>
              <td>${item.fields["System.Id"]}</td>
              <td>${item.fields["System.Title"]}</td>
              <td>${item.fields["System.State"]}</td>
              <td>${
                item.fields["System.AssignedTo"]?.displayName || "Unassigned"
              }</td>
              <td>${new Date(
                item.fields["System.CreatedDate"]
              ).toLocaleString()}</td>
              <td>${new Date(
                item.fields["System.ChangedDate"]
              ).toLocaleString()}</td>
              <td>${item.fields["System.Tags"] || ""}</td>
            </tr>`
      }

      htmlReport += `
          </table>
        </body>
        </html>`

      const htmlFilePath = "work_item_report.html"
      await fs.writeFile(htmlFilePath, htmlReport)

      return {
        content: [
          {
            type: "text",
            text: `Successfully fetched ${workItemsDetails.count} issues. Data saved to ${jsonFilePath} and HTML report generated at ${htmlFilePath}.`,
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
  "postReportToTeams",
  "Posts the work_item_report.html file as a formatted table to a Microsoft Teams channel.",
  {},
  {
    title: "Post Report to Teams",
    readOnlyHint: true,
    destructiveHint: false,
    idempotentHint: false,
    openWorldHint: true,
  },
  async () => {
    const teamsWebhookUrl =
      "https://centricconsultingllc.webhook.office.com/webhookb2/03340315-46bc-4fca-8d7d-575d3c54e7cc@d6f8cc30-debb-41a6-9c78-0516c185fa0d/IncomingWebhook/a7006be453ab46cfab41aa9c04adb08d/866683d8-0619-473b-8eaf-edd3c8cb36b9/V23tQ8mvzkNhQDHFF1Mg8ngIjrddKNuhtiye-_a1olYaU1";

    try {
      const html = await fs.readFile("work_item_report.html", "utf-8");

      // Load HTML into cheerio
      const $ = cheerio.load(html);

      // Extract headers
      const headers: string[] = [];
      $("table tr")
        .first()
        .find("th")
        .each((_, el) => {
          headers.push($(el).text().trim());
        });

      // Extract rows
      const rows: string[][] = [];
      $("table tr")
        .slice(1)
        .each((_, tr) => {
          const row: string[] = [];
          $(tr)
            .find("td")
            .each((_, td) => {
              row.push($(td).text().trim());
            });
          rows.push(row);
        });

      // Build Adaptive Card table-like body
      const cardBody: any[] = [];

      // Add header row
      cardBody.push({
        type: "ColumnSet",
        columns: headers.map((h) => ({
          type: "Column",
          width: "stretch",
          items: [
            {
              type: "TextBlock",
              text: `**${h}**`,
              wrap: true,
            },
          ],
        })),
      });

      // Add each data row
      rows.forEach((row) => {
        cardBody.push({
          type: "ColumnSet",
          columns: row.map((cell) => ({
            type: "Column",
            width: "stretch",
            items: [
              {
                type: "TextBlock",
                text: cell || "-",
                wrap: true,
              },
            ],
          })),
        });
      });

      // Adaptive Card payload
      const adaptiveCard = {
        type: "message",
        attachments: [
          {
            contentType: "application/vnd.microsoft.card.adaptive",
            content: {
              type: "AdaptiveCard",
              $schema: "http://adaptivecards.io/schemas/adaptive-card.json",
              version: "1.4",
              body: [
                {
                  type: "TextBlock",
                  text: "Azure DevOps Work Item Report",
                  weight: "Bolder",
                  size: "Medium",
                },
                {
                  type: "TextBlock",
                  text: `Generated on ${new Date().toLocaleString()}`,
                  isSubtle: true,
                  wrap: true,
                },
                ...cardBody,
              ],
            },
          },
        ],
      };

      const response = await fetch(teamsWebhookUrl, {
        method: "POST",
        headers: {
          "Content-Type": "application/json",
        },
        body: JSON.stringify(adaptiveCard),
      });

      if (!response.ok) {
        throw new Error(
          `Failed to post to Teams: ${response.status} ${await response.text()}`
        );
      }

      return {
        content: [
          {
            type: "text",
            text: "Report posted to Teams successfully in tabular format.",
          },
        ],
      };
    } catch (error) {
      return {
        content: [
          {
            type: "text",
            text: `Operation failed: ${error}`,
          },
        ],
      };
    }
  }
);



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
