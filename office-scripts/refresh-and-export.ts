type ExportRow={Year:number;Month:number;Branch:string;Group:string;Measure:string;Value:number;Target:number|string;Unit:string;SnapshotDate:string;RecordType:string;UpdatedAt:string};
type CellValue=string|number|boolean;
type RowRecord={[key:string]:CellValue};

function toRecord(headers:string[],values:CellValue[]):RowRecord{
  const record:RowRecord={};
  headers.forEach((header,index)=>{record[header]=values[index]??''});
  return record;
}

function main(workbook: ExcelScript.Workbook): {schema:number;isTestData:boolean;rows:ExportRow[]} {
  workbook.refreshAllDataConnections();
  workbook.getApplication().calculate(ExcelScript.CalculationType.full);
  const table=workbook.getTable('OpsDashboardExport');
  if(!table)throw new Error('Table OpsDashboardExport was not found.');
  const headers=table.getHeaderRowRange().getTexts()[0];
  const required=['Year','Month','Branch','Group','Measure','Value'];
  for(const name of required)if(!headers.includes(name))throw new Error(`Missing export column: ${name}`);
  const source:RowRecord[]=table.getRangeBetweenHeaderAndTotal().getValues().map((values:CellValue[])=>toRecord(headers,values));
  const allowedGroups=new Set(['Unit Savings','Monthly Unit Progress','Hours Worked Breakdown','Outlook','Agency vs Permanent','Budget','Gross Profit']);
  const weeklyGroups=new Set(['Unit Savings','Monthly Unit Progress','Hours Worked Breakdown']);
  const refreshedAt=new Date(),updatedAt=refreshedAt.toISOString(),snapshotDate=updatedAt.slice(0,10);
  const weekEndingDate=new Date(Date.UTC(refreshedAt.getUTCFullYear(),refreshedAt.getUTCMonth(),refreshedAt.getUTCDate()));
  const daysToSunday=(7-weekEndingDate.getUTCDay())%7;weekEndingDate.setUTCDate(weekEndingDate.getUTCDate()+daysToSunday);
  const weekEnding=weekEndingDate.toISOString().slice(0,10);
  const output:ExportRow[]=source.filter(r=>allowedGroups.has(String(r.Group))).map(r=>({
    Year:Number(r.Year),Month:Number(r.Month),Branch:String(r.Branch),Group:String(r.Group),Measure:String(r.Measure),
    Value:Number(r.Value),Target:r.Target===''?'':Number(r.Target),Unit:String(r.Unit||'number'),SnapshotDate:snapshotDate,RecordType:'Current',UpdatedAt:updatedAt
  }));
  if(!output.length)throw new Error('No valid aggregate dashboard rows were produced.');

  const latestPeriod=Math.max(...output.map(r=>r.Year*100+r.Month));
  const snapshot=output.filter(r=>weeklyGroups.has(r.Group)&&r.Year*100+r.Month===latestPeriod).map(r=>({...r,SnapshotDate:weekEnding,RecordType:'WeeklySnapshot'}));
  let history=workbook.getTable('OpsDashboardHistory');
  if(!history){
    const sheet=workbook.getWorksheet('Dashboard History')||workbook.addWorksheet('Dashboard History');
    sheet.getRange('A1:K1').setValues([['Year','Month','Branch','Group','Measure','Value','Target','Unit','SnapshotDate','RecordType','UpdatedAt']]);
    history=sheet.addTable('A1:K1',true);history.setName('OpsDashboardHistory');
  }
  const historyHeaders=history.getHeaderRowRange().getTexts()[0];
  const existing:RowRecord[]=history.getRowCount()?history.getRangeBetweenHeaderAndTotal().getValues().map((values:CellValue[])=>toRecord(historyHeaders,values)):[];
  const key=(r:RowRecord)=>[r.SnapshotDate,r.Year,r.Month,r.Branch,r.Group,r.Measure].join('|');
  const existingIndexes=new Map(existing.map((r,i)=>[key(r),i]));
  const historyBody=history.getRowCount()?history.getRangeBetweenHeaderAndTotal():null;
  const additions:(string|number|boolean)[][]=[];
  for(const r of snapshot){
    const values:CellValue[]=historyHeaders.map(h=>(r as unknown as RowRecord)[h]??'');
    const index=existingIndexes.get(key(r));
    if(index===undefined)additions.push(values);else historyBody?.getRow(index).setValues([values]);
  }
  if(additions.length)history.addRows(-1,additions);
  const cutoff=new Date(refreshedAt);cutoff.setUTCFullYear(cutoff.getUTCFullYear()-2);
  const retained:ExportRow[]=[...existing,...snapshot].filter(r=>String(r.RecordType)==='WeeklySnapshot'&&new Date(`${String(r.SnapshotDate).slice(0,10)}T00:00:00Z`)>=cutoff).map(r=>({
    Year:Number(r.Year),Month:Number(r.Month),Branch:String(r.Branch),Group:String(r.Group),Measure:String(r.Measure),Value:Number(r.Value),
    Target:r.Target===''?'':Number(r.Target),Unit:String(r.Unit||'number'),SnapshotDate:String(r.SnapshotDate).slice(0,10),RecordType:'WeeklySnapshot',UpdatedAt:String(r.UpdatedAt||updatedAt)
  }));
  return {schema:2,isTestData:false,rows:[...output,...retained]};
}
