type ExportRow={Year:number;Month:number;Branch:string;Group:string;Measure:string;Value:number;Target:number|string;Unit:string;UpdatedAt:string};

function main(workbook: ExcelScript.Workbook): {schema:number;isTestData:boolean;rows:ExportRow[]} {
  workbook.refreshAllDataConnections();
  workbook.getApplication().calculate(ExcelScript.CalculationType.full);
  const table=workbook.getTable('OpsDashboardExport');
  if(!table)throw new Error('Table OpsDashboardExport was not found.');
  const headers=table.getHeaderRowRange().getTexts()[0];
  const required=['Year','Month','Branch','Group','Measure','Value'];
  for(const name of required)if(!headers.includes(name))throw new Error(`Missing export column: ${name}`);
  const rows=table.getRangeBetweenHeaderAndTotal().getValues().map(values=>Object.fromEntries(headers.map((h,i)=>[h,values[i]])));
  const allowedGroups=new Set(['Unit Savings','Monthly Unit Progress','Hours Worked Breakdown','Outlook','Agency vs Permanent','Budget']);
  const output:ExportRow[]=rows.filter(r=>allowedGroups.has(String(r.Group))).map(r=>({
    Year:Number(r.Year),Month:Number(r.Month),Branch:String(r.Branch),Group:String(r.Group),Measure:String(r.Measure),
    Value:Number(r.Value),Target:r.Target===''?'':Number(r.Target),Unit:String(r.Unit||'number'),UpdatedAt:new Date().toISOString()
  }));
  if(!output.length)throw new Error('No valid aggregate dashboard rows were produced.');
  return {schema:1,isTestData:false,rows:output};
}
