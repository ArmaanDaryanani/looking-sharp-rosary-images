import fs from 'node:fs/promises';
import path from 'node:path';
import {execFileSync} from 'node:child_process';
import {SpreadsheetFile,FileBlob} from '@oai/artifact-tool';
// Usage: prepare SOURCE_DIR HOST_REPO OUTPUT_DIR, or finalize INPUT_XLSX HOST_REPO OUTPUT_DIR COMMIT
const [mode,input,repo,out,commit]=process.argv.slice(2);
const csv=rows=>rows.map(r=>r.map(v=>{const s=v==null?'':String(v);return /[",\r\n]/.test(s)?'"'+s.replaceAll('"','""')+'"':s;}).join(',')).join('\r\n')+'\r\n';
await fs.mkdir(out,{recursive:true});
const manifestPath=path.join(out,'complete-manifest.json');
if(mode==='prepare') {
 const labels=['Multicolor / Clear','Turquoise / Clear','Multicolor / Clear','White / Clear','Multicolor / Clear','Clear','Rose Pink / Clear','Black','Red / Clear','Amber','Rose Pink','Turquoise / Clear','Black / Clear','Red / Clear','Yellow / Clear','Lavender / Clear','Rose Pink / Clear','Black / Clear','Champagne / Clear','Multicolor / Clear','Multicolor / Clear','Multicolor / Clear','Rose Pink / Clear','Purple / Clear','Black / Clear','Green / Clear','Royal Blue / Clear','Turquoise / Clear','Red / Clear','Multicolor / Clear','Multicolor / Clear','Multicolor / Clear'];
 const colors=[['rose-pink','Rose Pink'],['red','Red'],['turquoise','Turquoise'],['denim-blue','Denim Blue'],['purple','Purple'],['orange','Orange'],['light-aqua','Light Aqua'],['black','Black'],['royal-blue','Royal Blue'],['green','Green'],['white','White']];
 const sources=labels.map((label,i)=>({source:String(i+1).padStart(2,'0')+'.png',label,pattern:'Original'}));
 for(const style of ['solid','alternating']) for(const [slug,label] of colors) sources.push({source:style+'-'+slug+'.png',label:label+(style==='alternating'?' / Clear':''),pattern:style==='solid'?'Solid':'Alternating'});
 const actual=(await fs.readdir(input)).filter(n=>/^\d{2}\.png$|^(solid|alternating)-.*\.png$/.test(n));
 if(actual.length!==54 || sources.some(s=>!actual.includes(s.source))) throw Error('Source coverage mismatch');
 await fs.mkdir(path.join(repo,'images'),{recursive:true});
 for(const [i,s] of sources.entries()) {
  s.id='2511BLT'+(12271+i); s.name=s.id; s.option=s.label.replaceAll(' / ','').replaceAll(' ',''); s.file=s.id+'-'+s.option+'.jpg';
  execFileSync('/usr/bin/sips',['-s','format','jpeg','-s','formatOptions','90','-z','1800','1800',path.join(input,s.source),'--out',path.join(repo,'images',s.file)],{stdio:'ignore'});
  s.bytes=(await fs.stat(path.join(repo,'images',s.file))).size;
  if(s.bytes<10000 || s.bytes>=3000000) throw Error('Image size invalid: '+s.file);
 }
 await fs.writeFile(manifestPath,JSON.stringify(sources,null,2));
 console.log(JSON.stringify({sources:54,first:sources[0].id,last:sources.at(-1).id,maxBytes:Math.max(...sources.map(s=>s.bytes))}));
} else if(mode==='finalize') {
 const sources=JSON.parse(await fs.readFile(manifestPath,'utf8'));
 const wb=await SpreadsheetFile.importXlsx(await FileBlob.load(input)); const sh=wb.worksheets.getItem('Products');
 const old=sh.getRange('A1:N23').values;
 const pre=await wb.render({sheetName:'Products',range:'H1:K4',scale:1.5}); await fs.writeFile(path.join(out,'before-complete.png'),new Uint8Array(await pre.arrayBuffer()));
 const checks=[];
 for(const s of sources) {
  s.url='https://raw.githubusercontent.com/ArmaanDaryanani/looking-sharp-rosary-images/'+commit+'/images/'+s.file;
  const res=await fetch(s.url); const b=Buffer.from(await res.arrayBuffer()); const expected=await fs.readFile(path.join(repo,'images',s.file));
  if(!res.ok || !res.headers.get('content-type')?.startsWith('image/') || !b.equals(expected)) throw Error('URL failed '+s.file);
  checks.push({file:s.file,status:res.status,bytes:b.length});
 }
 const rows=sources.map(s=>[s.id,s.name,s.option,s.label,'','Rhinestone rosary with '+s.label.toLowerCase()+' beads and a silver-tone crucifix.','Rosaries','Accessories',10.75,55,3,null,null,s.url]);
 for(let r=24;r<=55;r++) sh.getRange('A'+r+':N'+r).copyFrom(sh.getRange('A2:N2'),'all');
 sh.getRange('A1:N55').values=[old[0],...rows];
 sh.getRange('A2:N55').format.rowHeight=44;
 sh.getRange('F2:F55').format.wrapText=true;
 sh.getRange('N2:N55').format.wrapText=true;
 sh.getRange('I2:K55').format.fill='#FFFFFF';
 sh.getRange('I2:J55').setNumberFormat('0.00'); sh.getRange('K2:K55').setNumberFormat('0');
 wb.recalculate();
 if(new Set(rows.map(r=>r[0])).size!==54 || rows.some(r=>r[8]!==10.75||r[9]!==55||r[10]!==3||!r[13])) throw Error('Import validation failed');
 console.log((await wb.inspect({kind:'table',range:'Products!G1:K4',include:'values',tableMaxRows:4,tableMaxCols:5})).ndjson);
 for(const [name,range] of [['complete-prices','G1:K5'],['complete-names','A1:D5'],['complete-last','A51:D55']]) {
  const image=await wb.render({sheetName:'Products',range,scale:1.5}); await fs.writeFile(path.join(out,name+'.png'),new Uint8Array(await image.arrayBuffer()));
 }
 await (await SpreadsheetFile.exportXlsx(wb)).save(path.join(out,'Rosary-Import-Draft.xlsx'));
 await fs.writeFile(path.join(out,'Rosary-Import-Draft.csv'),csv([old[0],...rows]));
 await fs.writeFile(path.join(out,'Image-Name-Map.csv'),csv([['Style ID','Pattern','Color','Original Image','Brandboom Image'],...sources.map(s=>[s.id,s.pattern,s.label,s.source,s.file])]));
 await fs.writeFile(path.join(out,'complete-validation.json'),JSON.stringify({rows:54,wholesale:10.75,retail:55,minimum:3,urls:checks},null,2));
 console.log(JSON.stringify({rows:rows.length,verifiedUrls:checks.length,output:out}));
} else throw Error('Unknown mode');
