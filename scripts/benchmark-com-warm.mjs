// Balanced warm-search diagnosis: the control runs identical baseline code
// in an independent worker. Two epochs recreate all workers and swap startup
// order. Every epoch uses all six execution orders twice (12 rounds).
import assert from 'node:assert/strict';
import { mkdir, readFile, writeFile } from 'node:fs/promises';
import { createHash } from 'node:crypto';
import { performance } from 'node:perf_hooks';
import { parseArgs } from 'node:util';
import { pathToFileURL } from 'node:url';
import path from 'node:path';
const { values } = parseArgs({ options: { baseline: { type: 'string' }, current: { type: 'string' },
  out: { type: 'string' }, before: { type: 'string' } } });
for (const name of ['baseline', 'current', 'out', 'before']) if (!values[name]) throw Error(`${name} required`);
const output = path.resolve(values.out);
await mkdir(output);
const roots = { baseline: path.resolve(values.baseline), current: path.resolve(values.current), control: path.resolve(values.baseline) };
const names = Object.keys(roots);
const api = {};
const orders = [['baseline','current','control'], ['current','control','baseline'], ['control','baseline','current'],
  ['control','current','baseline'], ['current','baseline','control'], ['baseline','control','current']];
const params = { folder: 'inbox', before: values.before, query: 'the', search_body: true, scan_limit: 50, limit: 200 };
const report = { started_at: new Date().toISOString(), versions: {}, samples: [], method:
  'English warm search; same fixed before/scan 50/limit 200; 2 fresh-worker epochs, 12 balanced-order rounds each; independent identical-code baseline control; all results and paging compared.' };
const clean = item => Object.fromEntries(['id','subject','sender','received_at','unread','size_bytes','attachment_count','download_state','body_snippet'].map(key => [key,item[key]]));
const median = a => { const s = [...a].sort((x,y)=>x-y); return (s[Math.floor((s.length-1)/2)]+s[Math.floor(s.length/2)])/2; };
const save = () => writeFile(path.join(output,'results.json'), JSON.stringify(report,null,2));
try {
  for (const [name,root] of Object.entries(roots)) {
    const moduleAt = file => import(pathToFileURL(path.join(root,file)).href);
    api[name] = { ...await moduleAt('src/bridge.mjs'), ...await moduleAt('src/tools.mjs'), ...await moduleAt('src/windows-fulltext-index.mjs') };
    const hashes = {};
    for (const file of ['powershell/worker.ps1','src/windows-fulltext.mjs','src/windows-fulltext-index.mjs']) {
      hashes[file] = createHash('sha256').update(await readFile(path.join(root,file))).digest('hex');
    }
    report.versions[name] = { root, source_sha256: hashes };
  }
  for (let epoch=0;epoch<2;epoch++) {
    const entries = {};
    try {
      for (const name of epoch ? [...names].reverse() : names) {
        const index = new api[name].WindowsFulltextIndex({ directory: path.join(output, `${epoch+1}-${name}-index`) });
        const bridge = new api[name].OutlookBridge({ fulltextIndex:index });
        const entry = { index, bridge, stages:{} }; entries[name] = entry;
        const direct = bridge.call.bind(bridge);
        bridge.call = async (method,...args) => {
          const start=performance.now();
          try { return await direct(method,...args); }
          finally { (entry.stages[method]??=[]).push(performance.now()-start); }
        };
        await bridge.call('health', {});
        for(let i=0;i<2;i++) await bridge.call('search', api[name].validateCall('search',params));
      }
      for(let round=0;round<12;round++) {
        const order=orders[(round+epoch*3)%orders.length];
        const sample={ epoch:epoch+1, round:round+1, order, versions:{} };
        let reference;
        for(const name of order) {
          const entry=entries[name]; entry.stages={};
          const start=performance.now();
          const result=await entry.bridge.call('search',api[name].validateCall('search',params));
          const ms=performance.now()-start;
          assert.ok(!result.retry_required && !result.body_search_incomplete && result.fulltext.complete);
          assert.equal(result.fulltext.updated,0);
          if (result.fulltext.engine === 'direct-scan') {
            assert.equal(result.fulltext.indexed,0);
            assert.equal(result.fulltext.evaluated,result.fulltext.candidates);
            assert.equal(result.fulltext.index_state,'disabled');
          } else assert.equal(result.fulltext.reused,result.fulltext.candidates);
          const actual={ messages:result.messages.map(clean), scanned:result.scanned, next_offset:result.next_offset, complete:result.complete };
          if(reference) assert.deepEqual(actual,reference); else reference=actual;
          sample.versions[name]={ms,stages_ms:structuredClone(entry.stages),count:result.messages.length};
        }
        sample.results_equal=true; report.samples.push(sample); await save();
        console.log(JSON.stringify({epoch:epoch+1,round:round+1,ms:Object.fromEntries(names.map(n=>[n,sample.versions[n].ms]))}));
      }
    } finally { for(const {bridge,index} of Object.values(entries)) { bridge.close(); index.close(); } }
  }
  report.summary=Object.fromEntries([0,1,2].map(epoch=>{
    const samples=report.samples.filter(s=>!epoch || s.epoch===epoch);
    const medians=Object.fromEntries(names.map(n=>[n,median(samples.map(s=>s.versions[n].ms))]));
    const paired=Object.fromEntries(['current','control'].map(n=>[n,{
      median_percent_difference:median(samples.map(s=>(s.versions[n].ms/s.versions.baseline.ms-1)*100)),
      faster_rounds:samples.filter(s=>s.versions[n].ms<s.versions.baseline.ms).length,
      total_rounds:samples.length
    }]));
    return [epoch || 'all',{median_ms:medians,paired}];
  }));
  console.log(JSON.stringify(report.summary));
} catch(error) { report.error={code:error.code??'BENCHMARK_FAILED'}; process.exitCode=1; }
finally { report.finished_at=new Date().toISOString(); await save(); }
