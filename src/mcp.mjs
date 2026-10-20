import { McpServer } from '@modelcontextprotocol/server';
import { definitions } from './tools.mjs';

export function createMcp(call, { platform = process.platform } = {}) {
  const server = new McpServer({ name: 'outlook-local', version: '0.3.0' });
  const cacheWrites = platform !== 'win32';
  for (const [name, definition] of Object.entries(definitions)) {
    server.registerTool(`outlook_${name}`, {
      description: `${definition.description} Original mail is never modified.${cacheWrites ? ' Private cache files may be refreshed and pruned.' : ''}`,
      inputSchema: definition.schema,
      annotations: { readOnlyHint: !cacheWrites, destructiveHint: cacheWrites, idempotentHint: false, openWorldHint: false }
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
