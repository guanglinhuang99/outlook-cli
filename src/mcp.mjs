import { McpServer } from '@modelcontextprotocol/server';
import { definitionsForPlatform, schemaForPlatform } from './tools.mjs';
import { workflowToolsForPlatform } from './workflow-tools.mjs';

// Every tool only reads mail. Private index and cache upkeep does not change the user's
// mailbox or files, so tools are annotated read-only and clients need not prompt for them.
const annotations = { readOnlyHint: true, destructiveHint: false, idempotentHint: false, openWorldHint: false };

function reply(data) {
  return { content: [{ type: 'text', text: JSON.stringify(data) }], structuredContent: data };
}

function failure(error) {
  return { isError: true, content: [{ type: 'text', text: JSON.stringify({ error: { code: error.code ?? 'OUTLOOK_ERROR', message: error.message } }) }] };
}

export function toolNamesForPlatform(platform = process.platform) {
  return [...definitionsForPlatform(platform), ...workflowToolsForPlatform(platform)].map(([name]) => `outlook_${name}`);
}

export function createMcp(call, { platform = process.platform, workflows } = {}) {
  const server = new McpServer({ name: 'outlook-local', version: '0.3.0' });
  const note = ' Original mail is never modified.';
  for (const [name, definition] of definitionsForPlatform(platform)) {
    server.registerTool(`outlook_${name}`, { description: definition.description + note, inputSchema: definition.schema, annotations },
      async args => { try { return reply(await call(name, args)); } catch (error) { return failure(error); } });
  }
  if (workflows) {
    for (const [name, tool] of workflowToolsForPlatform(platform)) {
      server.registerTool(`outlook_${name}`, { description: tool.description + note, inputSchema: schemaForPlatform(tool.schema, platform), annotations },
        async args => { try { return reply(await tool.run(args, workflows)); } catch (error) { return failure(error); } });
    }
  }
  return server;
}
