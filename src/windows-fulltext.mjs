// Shared orchestration; each adapter supplies fresh, filtered metadata before
// the private index is consulted. Mailbox offsets are live, not snapshots.
export async function searchWindowsFulltext(bridge, params, index) {
  const deadline = Date.now() + 30000;
  const remaining = () => Math.max(1, deadline - Date.now());
  const page = await bridge.call('fulltext_candidates', {
    ...params, limit: 50, scan_limit: Math.min(params.scan_limit ?? 500, 50)
  }, { timeoutMs: remaining() });
  const candidates = page.messages.map(item => ({
    id: item.id, signature: item._fulltext_signature ?? null, guard: item._fulltext_guard ?? null
  }));
  const evaluated = await index.evaluate(candidates, {
    query: params.query ?? '', deadline,
    loadBody: (candidate, timeoutMs) => bridge.call('fulltext_read', {
      id: candidate.id, folder: page.folder_id ?? params.folder,
      signature: candidate.signature, guard: candidate.guard, validation: page.validation
    }, { timeoutMs: Math.min(timeoutMs, remaining()) })
  });
  // Source data can change even when all bodies were cache hits. Adapters
  // recheck the store generation (offline) or message guards (COM).
  if (page.validation) await bridge.call('fulltext_validate', {
    folder: page.folder_id ?? params.folder, validation: page.validation
  }, { timeoutMs: remaining() });
  const query = String(params.query ?? '').toLowerCase();
  const matches = [];
  let next = page.next_offset;
  for (const item of page.messages) {
    const metadataMatch = !query || [item.subject, item.sender?.name, item.sender?.address, item.sender?.smtp_address]
      .some(value => String(value ?? '').toLowerCase().includes(query));
    if (!metadataMatch && !evaluated.hits.has(item.id)) continue;
    const { _fulltext_signature, _fulltext_offset, _fulltext_guard, ...metadata } = item;
    if (evaluated.hits.has(item.id)) metadata.body_snippet = evaluated.hits.get(item.id);
    matches.push(metadata);
    if (matches.length >= (params.limit ?? 25)) {
      // Remaining candidates may also match. Resume after this actual item,
      // not after the entire indexed window.
      if (item !== page.messages.at(-1)) next = _fulltext_offset;
      break;
    }
  }
  const retry = evaluated.retry_required;
  const { validation, folder_id, ...publicPage } = page;
  return {
    ...publicPage, messages: matches, scan_limit_applied: Math.min(params.scan_limit ?? 500, 50),
    next_offset: retry ? (params.offset ?? 0) : next,
    complete: !retry && next === null,
    body_search_incomplete: evaluated.body_search_incomplete,
    retry_required: retry, fulltext: evaluated.fulltext,
    date_basis: 'received_at'
  };
}
