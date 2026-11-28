import { z } from 'zod';
import { readAttachmentText } from './attachment-workflows.mjs';
import { runReadParity } from './read-parity.mjs';

// MCP tools built from client-side workflows over the core service operations. Each
// context supplies `call(method, params, options)`, `readState()` and the private
// `stateDir` for delta checkpoints, so the stdio adapter and the HTTP service share them.
const messageId = z.string().regex(/^[A-Za-z0-9_-]{8,16000}$/);
const day = z.iso.date();
const scope = {
  folder: z.string().min(1).max(16000).optional()
    .describe('inbox, sent, all, drafts, deleted, junk, outbox, or a folder ID from stores/folders. Defaults: conversation all, delta inbox.'),
  include_subfolders: z.boolean().optional().describe('Include descendant folders (default true).'),
  from_date: day.optional().describe('First received date, inclusive (Asia/Singapore).'),
  to_date: day.optional().describe('Last received date, inclusive (Asia/Singapore).'),
  scan_limit: z.number().int().min(1).max(5000).optional().describe('Messages to scan, default 5000.')
};

export const workflowTools = {
  attachment_text: {
    platforms: ['darwin', 'win32'],
    description: 'Extract readable text from one attachment (txt/csv/md/json/xml/yaml, PDF, DOCX, XLSX) instead of base64 bytes. Use the 1-based index from attachments. truncated=true when max_chars was reached.',
    schema: z.strictObject({
      id: messageId, index: z.number().int().min(1).max(10000),
      max_chars: z.number().int().min(1).max(100000).optional().describe('Default 12000.'),
      backend: z.enum(['auto', 'com', 'local', 'mac-cache']).optional()
    }),
    run: (args, context) => readAttachmentText(args, { call: context.call, readState: context.readState })
  },
  conversation: {
    platforms: ['darwin'],
    description: 'Return the thread of one message: earlier and later mails linked by verified Message-ID references. subject_fallback=true also matches the normalized subject and marks those matches. include_body reads cached bodies of up to 25 thread mails.',
    schema: z.strictObject({
      ...scope, id: messageId,
      limit: z.number().int().min(1).max(100).optional(),
      subject_fallback: z.boolean().optional(),
      include_body: z.boolean().optional(),
      max_chars: z.number().int().min(1).max(12000).optional()
    }),
    run: (args, context) => runReadParity('conversation', args, context)
  },
  delta_bundle: {
    platforms: ['darwin'],
    description: 'Return only mail added, changed or gone since a checkpoint. Omit since for a baseline (all current mail as added). Pass the returned checkpoint as since next time with the same folder and date filters.',
    schema: z.strictObject({
      ...scope,
      since: z.string().regex(/^[a-f0-9]{64}$/).optional().describe('checkpoint from the previous delta_bundle result.')
    }),
    run: (args, context) => runReadParity('delta-bundle', args, context)
  }
};

export function workflowToolsForPlatform(platform = process.platform) {
  return Object.entries(workflowTools).filter(([, tool]) => tool.platforms.includes(platform));
}
