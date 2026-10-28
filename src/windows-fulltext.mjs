// Shared orchestration; each adapter supplies fresh, filtered metadata before
// the private index is consulted. Mailbox offsets are live, not snapshots.
export async function searchWindowsFulltext(bridge, params, index, { inlineColdRead = false, directScan = false, deferredIndex = false } = {}) {
  if (!String(params.query ?? '').trim()) return bridge.call('search', { ...params, search_body: false });
  const deadline = Date.now() + 30000;
  // Leave five seconds for the final source guard check even when building
  // the index consumes its entire allowance.
  const indexDeadline = deadline - 5000;
  const remaining = () => Math.max(1, deadline - Date.now());
  // A new COM cache starts with a memory-only scan. On the next query,
  // private cache preparation overlaps the fresh candidate scan; no body is
  // persisted before preparation succeeds. All preparation stays timed.
  const directBudget = directScan ? index.directReadBudget() : inlineColdRead && deferredIndex ? await index.beginDirectRead?.() : null;
  const deferredReady = inlineColdRead && !directBudget && deferredIndex ? index.prepareDeferredCache?.() : null;
  const cold = inlineColdRead && !directBudget && !deferredReady && index.beginColdRead ? await index.beginColdRead() : null;
  const coldBudget = directBudget ?? cold?.budget ?? (inlineColdRead && !index.beginColdRead ? await index.prepareColdRead?.() : null);
  const [page] = await Promise.all([bridge.call('fulltext_candidates', {
    ...params, limit: 50, scan_limit: Math.min(params.scan_limit ?? 500, 50), defer_metadata: true,
    ...(coldBudget ? { inline_body_budget: { ...coldBudget, budget_ms: Math.max(1, indexDeadline - Date.now() - 100) } } : {}),
    ...(directBudget ? { direct_scan: true, ...(directScan ? { direct_match_query: String(params.query).toLowerCase() } : {}) } : {})
  }, { timeoutMs: Math.max(1, indexDeadline - Date.now()) }), deferredReady ?? cold?.ready]);
  const candidates = page.messages.map(item => ({
    id: item.id, signature: item._fulltext_signature ?? null, guard: item._fulltext_guard ?? null
  }));
  let inlineBodies;
  if (coldBudget && page.inline_bodies !== undefined) {
    const ids = new Set(candidates.map(item => item.id));
    if (!Array.isArray(page.inline_bodies) || page.inline_bodies.length !== candidates.length
        || page.inline_bodies.some(item => !item || !ids.delete(item.id))) {
      throw Object.assign(new Error('Invalid inline body batch; retry the search'), { code: 'FULLTEXT_SOURCE_CHANGED' });
    }
    inlineBodies = new Map(page.inline_bodies.map(item => [item.id, item]));
  }
  if (directBudget && !inlineBodies) {
    throw Object.assign(new Error('Direct scan body batch is unavailable; retry the search'), { code: 'FULLTEXT_SOURCE_CHANGED' });
  }
  const direct = Boolean(directBudget);
  if (direct) {
    const proof = page.validation?.candidates;
    if (!Array.isArray(proof) || proof.length !== candidates.length || candidates.some((c, i) =>
      proof[i]?.id !== c.id || (proof[i]?.signature ?? null) !== c.signature || (proof[i]?.guard ?? null) !== c.guard)) {
      throw Object.assign(new Error('Direct scan has no matching source proof; retry the search'), { code: 'FULLTEXT_SOURCE_CHANGED' });
    }
  }
  const staged = direct ? new Map() : index.takeDeferredBodies?.(candidates) ?? new Map();
  let stagedReused = 0;
  const evaluated = direct ? index.evaluateDirect(candidates, inlineBodies, { query: params.query ?? '', deadline: indexDeadline, retainBodies: !directScan }) : await index.evaluate(candidates, {
    query: params.query ?? '', deadline: indexDeadline,
    loadBodies: inlineBodies ? async missing => missing.map(item => inlineBodies.get(item.id))
      : page.fulltext_capabilities?.batch_read ? async (missing, timeoutMs, maxBytes, maxBodyBytes) => {
      const toRead = missing.filter(item => !staged.has(item.id));
      stagedReused += missing.length - toRead.length;
      const result = toRead.length ? await bridge.call('fulltext_read_batch', {
        folder: page.folder_id ?? params.folder, candidates: toRead,
        budget_ms: Math.max(1, Math.min(30000, timeoutMs - 100)),
        max_bytes: Math.min(maxBytes, 128 * 1024 * 1024), max_body_bytes: Math.min(maxBodyBytes, 128 * 1024 * 1024)
      }, { timeoutMs: Math.min(timeoutMs, remaining()) }) : { bodies: [] };
      return [...missing.filter(item => staged.has(item.id)).map(item => staged.get(item.id)), ...result.bodies];
    } : undefined,
    loadBody: (candidate, timeoutMs) => {
      if (staged.has(candidate.id)) { stagedReused++; return staged.get(candidate.id); }
      return bridge.call('fulltext_read', {
      id: candidate.id, folder: page.folder_id ?? params.folder,
      signature: candidate.signature, guard: candidate.guard, validation: page.validation
      }, { timeoutMs: Math.min(timeoutMs, remaining()) });
    }
  });
  if (!direct && stagedReused) evaluated.fulltext.staged_reused = stagedReused;
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
  // Validate every candidate, including non-hits, against its original guard.
  // COM can fill expensive attachment counts only for actual results here.
  if (page.validation) {
    const deferred = page.fulltext_capabilities?.deferred_metadata;
    const validated = await bridge.call('fulltext_validate', {
      folder: page.folder_id ?? params.folder, validation: page.validation,
      ...(deferred ? { metadata_ids: matches.map(item => item.id) } : {})
    }, { timeoutMs: remaining() });
    if (direct && validated.valid !== true) {
      throw Object.assign(new Error('Direct scan source validation failed; retry the search'), { code: 'FULLTEXT_SOURCE_CHANGED' });
    }
    if (deferred) {
      const metadata = new Map((validated.metadata ?? []).map(item => [item.id, item]));
      for (const item of matches) {
        const count = metadata.get(item.id)?.attachment_count;
        if (!Number.isSafeInteger(count) || count < 0) {
          throw Object.assign(new Error('Validated result metadata is unavailable; retry the search'), { code: 'FULLTEXT_SOURCE_CHANGED' });
        }
        item.attachment_count = count;
      }
    }
  }
  if (direct && !directScan) index.stageValidatedBodies(evaluated.accepted);
  const retry = evaluated.retry_required;
  const { validation, folder_id, fulltext_capabilities, inline_bodies, ...publicPage } = page;
  return {
    ...publicPage, messages: matches, scan_limit_applied: Math.min(params.scan_limit ?? 500, 50),
    next_offset: retry ? (params.offset ?? 0) : next,
    complete: !retry && next === null,
    body_search_incomplete: evaluated.body_search_incomplete,
    retry_required: retry, fulltext: evaluated.fulltext,
    date_basis: 'received_at'
  };
}
