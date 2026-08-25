import fs from 'node:fs';

const groups=new Set(['Unit Savings','Monthly Unit Progress','Hours Worked Breakdown','Outlook','Agency vs Permanent','Budget','Gross Profit']);
const allowed=new Set(['Year','Month','Branch','Group','Measure','Value','Target','Unit','SnapshotDate','UpdatedAt']);
const payload=JSON.parse(process.env.OPS_PAYLOAD||'{}');
if(!Array.isArray(payload.rows))throw Error('client_payload.rows must be an array');
if(payload.rows.length>10000)throw Error('Aggregate payload is unexpectedly large');
for(const [i,row] of payload.rows.entries()){
  for(const key of Object.keys(row))if(!allowed.has(key))throw Error(`Row ${i+1} contains disallowed field: ${key}`);
  for(const key of ['Year','Month','Branch','Group','Measure','Value'])if(row[key]===undefined||row[key]===null||row[key]==='')throw Error(`Row ${i+1} is missing ${key}`);
  if(!groups.has(row.Group))throw Error(`Row ${i+1} has an unsupported group`);
  if(!Number.isFinite(Number(row.Value)))throw Error(`Row ${i+1} has a non-numeric value`);
}
const existing=JSON.parse(fs.readFileSync('dashboard-data.json','utf8'));
const incoming=payload.rows.map(row=>({...row,SnapshotDate:String(row.SnapshotDate||row.UpdatedAt||new Date().toISOString()).slice(0,10)}));
const key=row=>[row.SnapshotDate,row.Year,row.Month,row.Branch,row.Group,row.Measure].join('|');
const merged=new Map((existing.isTestData?[]:existing.rows||[]).map(row=>[key(row),row]));
for(const row of incoming)merged.set(key(row),row);
const cutoff=new Date();cutoff.setUTCFullYear(cutoff.getUTCFullYear()-2);
const rows=[...merged.values()].filter(row=>!row.SnapshotDate||new Date(`${row.SnapshotDate}T00:00:00Z`)>=cutoff);
fs.writeFileSync('dashboard-data.json',JSON.stringify({schema:2,isTestData:Boolean(payload.isTestData),rows},null,2)+'\n');
