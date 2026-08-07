import fs from 'node:fs';

const groups=new Set(['Unit Savings','Monthly Unit Progress','Hours Worked Breakdown','Outlook','Agency vs Permanent','Budget']);
const allowed=new Set(['Year','Month','Branch','Group','Measure','Value','Target','Unit','UpdatedAt']);
const payload=JSON.parse(process.env.OPS_PAYLOAD||'{}');
if(!Array.isArray(payload.rows))throw Error('client_payload.rows must be an array');
if(payload.rows.length>10000)throw Error('Aggregate payload is unexpectedly large');
for(const [i,row] of payload.rows.entries()){
  for(const key of Object.keys(row))if(!allowed.has(key))throw Error(`Row ${i+1} contains disallowed field: ${key}`);
  for(const key of ['Year','Month','Branch','Group','Measure','Value'])if(row[key]===undefined||row[key]===null||row[key]==='')throw Error(`Row ${i+1} is missing ${key}`);
  if(!groups.has(row.Group))throw Error(`Row ${i+1} has an unsupported group`);
  if(!Number.isFinite(Number(row.Value)))throw Error(`Row ${i+1} has a non-numeric value`);
}
fs.writeFileSync('dashboard-data.json',JSON.stringify({schema:1,isTestData:Boolean(payload.isTestData),rows:payload.rows},null,2)+'\n');
