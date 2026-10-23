// Core regression tests. Original exports are optional and never committed.
const assert=require('node:assert/strict');
const fs=require('node:fs');
const path=require('node:path');
const {execFileSync}=require('node:child_process');
global.JSZip=require('../vendor/jszip.min.js');
require('../converter.js');
const C=global.SkuConverter;
let checks=0;
const check=(value,message)=>{assert.ok(value,message);checks++;};
const rows=data=>data.map((values,i)=>({number:i+1,cells:values.map(text=>({text:String(text),type:'s'}))}));
const convert=(marketplace,brand,data)=>C.convert(rows(data),{marketplace,brand,country:'MY'});
async function main(){
  const sh=convert('Shopee','Vans',[['SKU','Variation ID','Product ID'],['001SKU','00456','00789'],['  ','2','123']]);
  check(sh.records[0][3]==='00789~00456','text ID leading zeros');
  check(sh.records[0][1]==='001SKU'&&sh.records[0][2]==='001SKU','both SKU columns');
  check(sh.records[1][1]==='0'&&sh.records[1][2]==='0'&&sh.blankSkus===1,'blank SKU to zero');
  check(sh.records[0][0]==='M02'&&sh.records[0][8]==='Active'&&sh.records[0][9]==='Create','fixed values');
  check(sh.records[0].slice(4,8).every(v=>v===''),'optional fields blank');
  const la=convert('Lazada','Vans',[['SellerSKU','sku.skuId'],['Optional','Optional'],['description','description'],['rules',''],['SKU123','123']]);
  check(la.records.length===1&&la.skippedInstructions===3&&la.records[0][0]==='M06','Lazada metadata and code');
  check(convert('Lazada','Columbia',[['sku.skuId','SellerSKU'],['123','ABC']]).records.length===1,'Lazada without metadata');
  const za=convert('Zalora','Timberland',[['CatalogProductId','SellerSku','Brand'],['123','ABC','Timberland']]);
  check(za.records[0][0]==='M16'&&za.records[0][3]==='123','Zalora mapping');
  const dup=convert('Shopee','Vans',[['Product ID','Variation ID','SKU'],['1','2','A'],['1','2','A'],['','','']]);
  check(dup.duplicateMappings===1&&dup.records.length===2&&dup.emptyRows===1,'duplicates kept; empty rows skipped');
  const failureCases=[
    ()=>convert('Shopee','Vans',[['SKU','Variation ID','Product ID'],['A','','123']]),
    ()=>convert('Lazada','Vans',[['SellerSKU','sku.skuId'],['A','']]),
    ()=>convert('Lazada','Vans',[['wrong'],['data']]),
    ()=>convert('Zalora','Vans',[['CatalogProductId','SellerSku','Brand'],['1','A','Timberland']]),
    ()=>convert('Shopee','Vans',[['SKU','Variation ID','Product ID'],['A','x','1']]),
    ()=>convert('Shopee','Vans',[['SKU','SKU','Variation ID','Product ID'],['A','A','2','1']]),
  ];
  for(const fn of failureCases){assert.throws(fn);checks++;}
  for(const cell of [{text:'1234567890123456',type:'n'},{text:'1.23E+14',type:'n'},{text:'123',formula:true,type:'n'},{text:'#N/A',type:'e'}]){
    const input=rows([['Product ID','Variation ID','SKU'],['1','2','ABC']]);input[1].cells[0]=cell;
    assert.throws(()=>C.convert(input,{marketplace:'Shopee',brand:'Vans',country:'MY'}));checks++;
  }
  check(C.CHANNELS.MY.Shopee.Columbia==='M01'&&C.CHANNELS.MY.Lazada.Timberland==='M05'&&C.CHANNELS.MY.Zalora.Vans==='M19','channel codes');
  const template=fs.readFileSync(path.join(__dirname,'../template.xls'));
  const base=await JSZip.loadAsync(template);
  const baseSheet=await base.file('xl/worksheets/sheet1.xml').async('string');
  const headerRows=text=>text.match(/<row\b[^>]*r="[12]"[^>]*>[\s\S]*?<\/row>/g);
  for(const size of [1,5198,5199,12000]){
    const result={selection:{marketplace:'Shopee',brand:'Vans',country:'SG'},records:Array.from({length:size},(_,i)=>['M02','001SKU'+i,'001SKU'+i,'123~'+i,'','','','','Active','Create'])};
    const batches=C.batches(result);
    check(batches.reduce((n,b)=>n+b.records.length,0)===size&&batches.every(b=>b.records.length<=5198),'split complete: '+size);
    if(size===12000)check(JSON.stringify(batches.map(b=>b.records.length))==='[5198,5198,1604]','12000 exact split');
    for(const batch of batches){
      const bytes=await C.exportBatch(template,batch.records);const out=await JSZip.loadAsync(bytes);
      const sheet=await out.file('xl/worksheets/sheet1.xml').async('string');
      check((sheet.match(/<row\b/g)||[]).length===batch.records.length+2,'export row limit');
      assert.deepEqual(headerRows(sheet),headerRows(baseSheet));
      check(sheet.includes(`A1:J${batch.records.length+2}`),'used range');
      for(const [name,entry]of Object.entries(base.files))if(!entry.dir&&name!=='xl/worksheets/sheet1.xml'&&name!=='xl/sharedStrings.xml')assert.deepEqual(await out.file(name).async('uint8array'),await entry.async('uint8array'),'unchanged template part '+name);
      // Independently parse exported XML with Python's XML parser.
      const strings=await out.file('xl/sharedStrings.xml').async('string');
      const summary=JSON.parse(execFileSync('python3',['-c',"import sys,json,xml.etree.ElementTree as E;data=json.loads(sys.stdin.read());r=E.fromstring(data['sheet']);n={'m':'http://schemas.openxmlformats.org/spreadsheetml/2006/main'};ss=[''.join(i.itertext()) for i in E.fromstring(data['strings'])];rows=r.findall('m:sheetData/m:row',n);print(json.dumps({'count':len(rows),'first':ss[int(rows[2].findall('m:c',n)[1].find('m:v',n).text)],'last':ss[int(rows[-1].findall('m:c',n)[1].find('m:v',n).text)]}))"],{input:JSON.stringify({sheet,strings}),encoding:'utf8'}));
      check(summary.count===batch.records.length+2&&summary.first===batch.records[0][1]&&summary.last===batch.records.at(-1)[1],'independent XML parse / first and last row');
    }
  }
  const escaped=await JSZip.loadAsync(await C.exportBatch(template,[['M02','=VALUE<&"','00123','1~2','','','','','Active','Create']]));
  const escapedSheet=await escaped.file('xl/worksheets/sheet1.xml').async('string');
  const escapedStrings=await escaped.file('xl/sharedStrings.xml').async('string');
  check(escapedStrings.includes('=VALUE&lt;&amp;&quot;')&&!escapedSheet.includes('<f>'),'XML escaping and formula-like text');
  check(!escapedSheet.includes('inlineStr')&&/<c r="A3" t="s"><v>\d+<\/v><\/c>/.test(escapedSheet),'A3 uses template-compatible shared string');
  const sampleDir=process.env.SAMPLE_DIR;
  if(sampleDir){
    for(const [market,brand,count,blank]of [['Shopee','Vans',5,1],['Lazada','Vans',6,0],['Zalora','Timberland',12,0]]){
      const filename=path.join(sampleDir,market+'LayoutExample.xlsx');
      const source=JSON.parse(execFileSync('python3',['-c',`import sys,json,zipfile,xml.etree.ElementTree as E
n={'m':'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
with zipfile.ZipFile(sys.argv[1]) as z:
 ss=[]
 if 'xl/sharedStrings.xml' in z.namelist():ss=[''.join(t.text or '' for t in i.iter('{'+n['m']+'}t')) for i in E.fromstring(z.read('xl/sharedStrings.xml'))]
 result=[]
 for r in E.fromstring(z.read('xl/worksheets/sheet1.xml')).findall('m:sheetData/m:row',n):
  cells=[]
  for c in r:
   ref=c.get('r');letters=''.join(x for x in ref if x.isalpha());idx=0
   for x in letters:idx=idx*26+ord(x)-64
   while len(cells)<idx:cells.append(None)
   v=c.find('m:v',n);text=v.text if v is not None else '';typ=c.get('t')
   if typ=='s':text=ss[int(text)]
   elif typ=='inlineStr':text=''.join(t.text or '' for t in c.iter('{'+n['m']+'}t'))
   cells[idx-1]={'text':text,'type':typ,'formula':c.find('m:f',n)is not None}
  result.append({'number':int(r.get('r')),'cells':cells})
 print(json.dumps(result))`,filename],{encoding:'utf8'}));
      if(market==='Lazada'&&!source.some(row=>row.cells.some(cell=>cell?.text==='sku.skuId'))){
        assert.throws(()=>C.convert(source,{marketplace:market,brand,country:'MY'}),/sku.skuId/);checks++;continue;
      }
      const result=C.convert(source,{marketplace:market,brand,country:'MY'});
      check(result.records.length===count,market+' sample count');check(result.blankSkus===blank,market+' sample blank SKU');
      if(market==='Shopee')check(result.records[0][3]==='42611267358~265946333448','actual Shopee identifier');
    }
  }
  console.log(`PASS: ${checks} core checks. Original samples: ${sampleDir?'included':'not supplied'}. Browser/UI checks run separately.`);
}
main().catch(error=>{console.error(error);process.exitCode=1;});
