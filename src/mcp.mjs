import { McpServer } from '@modelcontextprotocol/server';
import { definitions } from './tools.mjs';

export function createMcp(call) {
  const server = new McpServer({ name: 'outlook-local', version: '0.2.1' });
  for (const [name, definition] of Object.entries(definitions)) {
    server.registerTool(`outlook_${name}`, {
      description: definition.description,
      inputSchema: definition.schema,
      annotations: { readOnlyHint: true, destructiveHint: false, idempotentHint: true, openWorldHint: false }
    }, async args => {
      try {
        const data = await call(name, args);
        return { content: [{ type: 'text', text: JSON.stringify(data) }], structuredContent: data };
      } catch (error) {
        return { isError: true, content: [{ type: 'text', text: JSON.stringify({ error: { code: error.code ?? 'OUTLOOK_ERROR', message: error.message } }) }] };
      }
    });
  }
  return server;
}
