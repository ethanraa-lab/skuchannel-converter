'use strict';
const $=id=>document.getElementById(id);
let uploads=[],result=null,templatePromise=null,busy=false;
const converter=globalThis.SkuConverter;
function message(text,error=false){$('message').textContent=text;$('message').className=error?'error':'';}
function reset(){result=null;$('results').hidden=true;$('empty').hidden=false;$('download-all').disabled=true;message('');}
const channels=Object.entries(converter.CHANNELS).flatMap(([marketplace,brands])=>
  Object.entries(brands).map(([brand,code])=>({marketplace,brand,code}))
).sort((a,b)=>a.code.localeCompare(b.code));
function selection(){
  const channel=channels.find(item=>item.code===$('channel').value);
  return {marketplace:channel.marketplace,brand:channel.brand,country:document.querySelector('input[name="country"]:checked').value};
}
function updateCode(){ $('channel-code').textContent=$('channel').value; }
function setBusy(value){
  busy=value;
  for(const id of ['channel','country','file']) $(id).disabled=value;
  $('convert').disabled=value||!uploads.length;
  $('download-all').disabled=value||!result;
  document.querySelectorAll('.batch button,.remove-file,.file-sheet').forEach(control=>control.disabled=value);
}
for(const id of ['channel','country']) $(id).addEventListener('change',()=>{updateCode();reset();});
function renderUploads(){
  $('file-name').textContent=uploads.length ? uploads.length+' file(s) selected' : 'No file selected';
  $('file-list').replaceChildren(...uploads.map((item,index)=>{
    const group=document.createElement('div');group.className='upload-group';
    const row=document.createElement('div');row.className='uploaded-file';
    const info=document.createElement('div');info.className='file-info';
    const name=document.createElement('strong');name.textContent=item.file.name;info.append(name);
    group.append(row);
    if(item.workbook.sheets.length>1){
      const worksheet=document.createElement('div');worksheet.className='worksheet-box';
      const label=document.createElement('label');label.textContent='Worksheet';
      const select=document.createElement('select');select.className='file-sheet';select.id='file-sheet-'+index;
      label.htmlFor=select.id;
      select.replaceChildren(...item.workbook.sheets.map((sheet,i)=>new Option(sheet.name.trim().toLowerCase()==='template'?'SKUChannel Template':sheet.name,String(i))));
      select.value=String(item.sheet);
      select.addEventListener('change',()=>{item.sheet=Number(select.value);reset();});
      worksheet.append(label,select);group.append(worksheet);
    }
    const remove=document.createElement('button');remove.type='button';remove.className='remove-file';
    remove.textContent='×';remove.setAttribute('aria-label','Remove '+item.file.name);remove.title='Remove file';
    remove.addEventListener('click',()=>{
      if(busy)return;
      uploads.splice(index,1);$('file').value='';reset();renderUploads();setBusy(false);
    });
    row.append(info,remove);return group;
  }));
}
$('file').addEventListener('change',async()=>{
  const files=Array.from($('file').files);$('file').value='';if(!files.length)return;
  reset();setBusy(true);const errors=[];
  try{
    for(const file of files){
      message('Reading '+file.name+'…');
      try{
        if(file.size>30*1024*1024)throw new Error('File exceeds 30 MB.');
        if(uploads.reduce((sum,item)=>sum+item.file.size,0)+file.size>90*1024*1024)throw new Error('Combined uploads exceed 90 MB. Convert these files first.');
        if(uploads.some(item=>item.file.name===file.name&&item.file.size===file.size&&item.file.lastModified===file.lastModified))throw new Error('This file is already in the list.');
        uploads.push({file,workbook:await converter.openWorkbook(await file.arrayBuffer()),sheet:0});
      }catch(error){errors.push(file.name+': '+error.message);}
    }
    renderUploads();message(errors.length?errors.join('\n'):'Files ready. All files must match the selected marketplace and brand.',errors.length>0);
  }finally{setBusy(false);}
});
$('convert').addEventListener('click',async()=>{
  if(busy||!uploads.length)return;reset();setBusy(true);message('Converting records…');
  try{
    await new Promise(resolve=>setTimeout(resolve,0));
    const selected=selection();
    const combined={records:[],blankSkus:0,duplicateMappings:0,emptyRows:0,skippedInstructions:0,selection:selected};
    const seen=new Set();
    for(const item of uploads){
      let converted;
      try{converted=converter.convert(await item.workbook.rows(item.sheet),selected);}
      catch(error){throw new Error(item.file.name+': '+error.message);}
      for(const key of ['blankSkus','emptyRows','skippedInstructions'])combined[key]+=converted[key];
      for(const record of converted.records){
        const key=JSON.stringify([record[1],record[3]]);
        if(seen.has(key))combined.duplicateMappings++;
        seen.add(key);combined.records.push(record);
      }
    }
    combined.batches=Math.ceil(combined.records.length/converter.LIMIT);
    result=combined;render();message('Conversion complete. Review the preview before downloading.');
  }catch(error){message(error.message,true);}finally{setBusy(false);}
});
function render(){
  $('results').hidden=false;$('empty').hidden=true;
  $('record-count').textContent=result.records.length.toLocaleString();$('batch-count').textContent=result.batches;$('blank-count').textContent=result.blankSkus;
  $('summary').textContent=`${result.selection.marketplace} ${result.selection.country} · ${result.selection.brand} · ${result.skippedInstructions} instruction rows skipped · ${result.emptyRows} empty rows skipped`;
  $('warnings').replaceChildren();
  const warnings=[];
  if(result.blankSkus)warnings.push(`${result.blankSkus} blank SKU(s) were replaced with 0 in both SKU columns. Check these records before importing into eRetail.`);
  if(result.duplicateMappings)warnings.push(`${result.duplicateMappings} repeated SKU/product mapping(s) retained. No records were deduplicated.`);
  for(const text of warnings){const p=document.createElement('p');p.className='warning';p.textContent=text;$('warnings').append(p);}
  $('preview').replaceChildren(...result.records.slice(0,20).map(record=>{
    const tr=document.createElement('tr');for(const i of [0,1,2,3,8,9]){const td=document.createElement('td');td.textContent=record[i];tr.append(td);}return tr;
  }));
  $('batches').replaceChildren(...converter.batches(result).map(batch=>{
    const div=document.createElement('div');div.className='batch';const span=document.createElement('span');span.textContent=batch.name;
    const small=document.createElement('small');small.textContent=`${batch.records.length.toLocaleString()} records + 2 header rows`;span.append(small);
    const button=document.createElement('button');button.textContent='Download';button.setAttribute('aria-label','Download '+batch.name);
    button.addEventListener('click',()=>downloadBatch(batch));div.append(span,button);return div;
  }));
}
async function template(){
  if(!templatePromise)templatePromise=fetch('template.xls').then(response=>{if(!response.ok)throw new Error('Could not load the eRetail template.');return response.arrayBuffer();}).catch(error=>{templatePromise=null;throw error;});
  return templatePromise;
}
function save(bytes,name,type){
  const url=URL.createObjectURL(new Blob([bytes],{type}));const a=document.createElement('a');a.href=url;a.download=name;document.body.append(a);a.click();a.remove();setTimeout(()=>URL.revokeObjectURL(url),10000);
}
async function downloadBatch(batch){
  if(busy)return;setBusy(true);message('Preparing '+batch.name+'…');
  try{save(await converter.exportBatch(await template(),batch.records),batch.name,'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');message('File prepared. Check your browser downloads.');}
  catch(error){message(error.message,true);}finally{setBusy(false);}
}
$('download-all').addEventListener('click',async()=>{
  if(busy||!result)return;setBusy(true);
  try{
    const zip=new JSZip(),source=await template();const files=converter.batches(result);
    for(let i=0;i<files.length;i++){message(`Preparing file ${i+1} of ${files.length}…`);zip.file(files[i].name,await converter.exportBatch(source,files[i].records));}
    const {marketplace,country,brand}=result.selection;
    save(await zip.generateAsync({type:'uint8array',compression:'STORE'}),`SkuChannel_${marketplace}_${country}_${brand}.zip`,'application/zip');message('ZIP prepared. Check your browser downloads.');
  }catch(error){message(error.message,true);}finally{setBusy(false);}
});
$('channel').replaceChildren(...channels.map(({marketplace,brand,code})=>new Option(`${code} - ${marketplace} ${brand}`,code)));
updateCode();
