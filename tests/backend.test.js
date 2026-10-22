// Backend tests: node tests/backend.test.js  (mocks SpreadsheetApp, Utilities, etc.; prints ok/FAIL per check)
const fs=require('fs');const crypto=require('crypto');
const code=fs.readFileSync(require('path').join(__dirname,'..','src','Code.gs'),'utf8');
function mkSheet(head,rows){const data=[head,...rows];const sh={data,
 getDataRange(){return{getValues:()=>data.map(r=>r.slice())}},getLastColumn(){return Math.max(...data.map(r=>r.length))},getLastRow(){return data.length},getMaxRows(){return 1000},
 getRange(r,c,nr,nc){nr=nr||1;nc=nc||1;return{getValues:()=>{const o=[];for(let i=0;i<nr;i++){const row=data[r-1+i]||[];o.push(Array.from({length:nc},(_,j)=>row[c-1+j]??''));}return o;},
   setValues(v){v.forEach((row,i)=>{data[r-1+i]=data[r-1+i]||[];row.forEach((x,j)=>data[r-1+i][c-1+j]=x);});return this},setValue:v=>{data[r-1]=data[r-1]||[];data[r-1][c-1]=v},setNumberFormat(){},setWrap(){return this},setFontSize(){return this},setFontWeight(){return this}}},
 appendRow(r){data.push(r)},clear(){data.length=0}};return sh;}
const sheets={
 'Setup':mkSheet(['old'],[['x']]),
 'Users':mkSheet(['Name','Email','Active','Admin'],[['Mark Ellis','m@h.com','Y','Y'],['Joe Site','','Y','N']]),
 'Items':mkSheet(['ItemID','Walk','Trade','Area','Hold','Item','Division','Order','Active'],[['t001',2,'Framing','backframing','HOLD','Backing installed','Div 06-1',1,'Y'],['t002',3,'Drywall','','','Screws mudded','Div 09',2,'Y'],['t003',4,'Paint','','','Paint done','Div 09',3,'Y']]),
 'Jobs':mkSheet(['JobID','Job Name','Address','Created By','Created At','Active'],[['J1','Old Job','1 St','Mark',new Date(),'Y']]),
 'Job Items':mkSheet(['JobID','ItemID','Type','Walk','Trade','Area','Hold','Item','Added By','Added At','Removed'],[['J1','X1','Item',1,'Plumbing','','HOLD','Job extra','Mark',new Date(),'']]),
 'Check Log':mkSheet(['JobID','ItemID','Walk','Status','Name','Timestamp'],[['J1','t001',2,'Done','Mark',new Date()]])};
const props={},cache={};
const ctx={SpreadsheetApp:{getActive:()=>({getSheetByName:n=>sheets[n]||null,insertSheet:n=>{sheets[n]=mkSheet([],[]);sheets[n].data.length=0;sheets[n].setFrozenRows=()=>{};return sheets[n];}})},
 PropertiesService:{getScriptProperties:()=>({getProperty:k=>props[k]||null,setProperty:(k,v)=>props[k]=v})},
 CacheService:{getScriptCache:()=>({get:k=>cache[k]??null,put:(k,v)=>cache[k]=v,remove:k=>delete cache[k]})},
 LockService:{getScriptLock:()=>({waitLock(){if(this.h)throw new Error('deadlock');this.h=1},releaseLock(){this.h=0}})},HtmlService:{},
 Utilities:{getUuid:()=>crypto.randomUUID(),computeHmacSha256Signature:(p,s)=>[...crypto.createHmac('sha256',s).update(p).digest()],
  base64EncodeWebSafe:x=>Buffer.from(typeof x==='string'?x:Buffer.from(x.map(b=>b&255))).toString('base64').replace(/\+/g,'-').replace(/\//g,'_'),
  base64DecodeWebSafe:s=>[...Buffer.from(s.replace(/-/g,'+').replace(/_/g,'/'),'base64')],newBlob:b=>({getDataAsString:()=>Buffer.from(b.map(x=>x&255)).toString()}),formatDate:d=>d.toISOString().slice(5,16)}};
const A=new Function(...Object.keys(ctx),code+';return {editItem,setup,login,whoAmI,getJobs,createJob,getJob,setCheck,addItems,removeItem,addNote};')(...Object.values(ctx));
const ok=(c,m)=>{if(!c){console.log('FAIL',m);process.exitCode=1}else console.log('ok',m)};
// login before PINs exist -> rejected, but upgrade runs
try{A.login('Mark Ellis','1');ok(false,'no pin')}catch(e){console.log('   msg:',e.message);ok(/not recognised/.test(e.message),'no PIN yet -> rejected')}
const U=sheets['Users'].data;ok(U[0].join()==='Name,Email,Active,Admin,PIN','Users: PIN column added');
const I=sheets['Items'].data;ok(I[0][1]==='Checklist'&&I[0][4]==='Critical','Items headers renamed');
ok(I[1][1]===4&&I[2][1]===5&&I[3][1]===6,'existing checklists shifted to 4/5/6');
ok(I[1][4]==='CRITICAL','HOLD -> CRITICAL');
ok(I.filter(r=>r[1]===1).length===4&&I.find(r=>r[0]==='S001')[5]==='Site signage','4 startup items added to checklist 1');
ok(sheets['Job Items'].data[1][3]===3&&sheets['Job Items'].data[0][3]==='Checklist','Job Items shifted');
ok(sheets['Check Log'].data[1][2]===4,'Check Log shifted');
ok(String(sheets['Setup'].data[0][0]).includes('SITE CHECKLISTS'),'Setup tab rewritten');
// idempotent
delete cache['schema'];A.setup();ok(I.filter(r=>r[1]===1).length===4&&I[1][1]===4,'second upgrade run changes nothing');
U[1][4]='0123';U[2][4]='5555';
let r=A.login('mark ellis','123');ok(r.admin,'numeric PIN with leading zero');
for(let i=0;i<5;i++)try{A.login('Joe Site','9')}catch(e){}
try{A.login('Joe Site','5555');ok(false,'lock')}catch(e){ok(/Too many/.test(e.message),'lockout')}
const t=r.token;let d=A.getJob(t,'J1');
console.log('   items',d.items.map(i=>i.id+':'+i.list).join(' '),'check',JSON.stringify(d.checks));ok(d.items.length===8&&d.checks['t001'].status==='Done','old job keeps its tick after upgrade');
ok(d.items.find(i=>i.id==='t001').critical&&d.items.find(i=>i.id==='X1').critical,'critical flags read');
const j=A.createJob(t,'Chen','12 Main');ok(A.getJob(t,j.id).items.filter(i=>i.list===1).length===4,'new job gets startup checklist');
ok(A.setCheck(t,j.id,'S001',1,'Done').name==='Mark Ellis','tick on checklist 1');
ok(A.addItems(t,j.id,'all',[{list:2,trade:'Demo',critical:true,text:'Hazmat report on site'}])===1,'admin adds to Demolition (inside lock, no deadlock)');
ok(A.getJob(t,j.id).items.find(i=>i.list===2).critical,'new item critical');
A.removeItem(t,j.id,'t002','m',5);ok(!A.getJob(t,j.id).items.some(i=>i.id==='t002'),'hide master item for job');
U[1][2]='N';try{A.getJobs(t);ok(false,'blk')}catch(e){ok(e.message==='AUTH','blocked user loses access')}

U[1][2]='Y';
try{A.editItem(r3t(),'J1','t003','m',{text:'x'});}catch(e){}
function r3t(){U[2][4]='5555';delete cache['fail:joe site'];return A.login('Joe Site','5555').token;}
try{A.editItem(r3t(),'J1','t003','m',{text:'New'});ok(false,'nonadmin edit')}catch(e){ok(/Only admins/.test(e.message),'non-admin cannot edit')}
const t2=A.login('Mark Ellis','0123').token;
let n=A.editItem(t2,'J1','t003','m',{text:'Paint touch-ups complete',trade:'Painting',area:'ceilings',critical:true});
ok(n===4,'admin edit changes 4 fields');
const it=sheets['Items'].data.find(r=>r[0]==='t003');ok(it[5]==='Paint touch-ups complete'&&it[2]==='Painting'&&it[3]==='ceilings'&&it[4]==='CRITICAL','master item updated in sheet');
const L=sheets['Edit Log'].data;ok(L[0][0]==='Timestamp'&&L.length===5&&L.some(r=>r[4]==='Wording'&&r[5]==='Paint done'&&r[6]==='Paint touch-ups complete'&&r[1]==='Mark Ellis'),'Edit Log records old -> new with name');
ok(A.editItem(t2,'J1','t003','m',{text:'Paint touch-ups complete',trade:'Painting',area:'ceilings',critical:true})===0,'no-op edit logs nothing');
ok(A.editItem(t2,'J1','X1','j',{text:'Job extra (edited)',trade:'Plumbing',area:'',critical:false})===2,'job-only item edit');
ok(sheets['Job Items'].data.find(r=>r[1]==='X1')[7]==='Job extra (edited)','job item updated');
try{A.editItem(t2,'J1','t003','m',{text:'  '});ok(false,'empty')}catch(e){ok(/cannot be empty/.test(e.message),'empty wording rejected')}
const d2=A.getJob(t2,'J1');ok(d2.checks['t001'].status==='Done','ticks unaffected by edits');
// notes (anyone signed in, running log per job)
ok(!sheets['Notes']&&JSON.stringify(A.getJob(t2,'J1').notes)==='{}','no Notes tab yet -> job loads with no notes');
const tj=r3t();const n1=A.addNote(tj,'J1','t001',4,'  Backing short at master bath  ');
ok(n1.name==='Joe Site'&&n1.text==='Backing short at master bath','non-admin can add a note (trimmed, with name)');
ok(sheets['Notes']&&sheets['Notes'].data[0].join()==='JobID,ItemID,Checklist,Note,Name,Timestamp','Notes tab created with headings');
A.addNote(t2,'J1','t001',4,'Fixed Oct 22');A.addNote(t2,j.id,'t001',4,'Other job note');
const nn=A.getJob(t2,'J1').notes['t001'];ok(nn.length===2&&nn[0].name==='Joe Site'&&nn[1].text==='Fixed Oct 22','notes kept as a running log, oldest first');
ok(A.getJob(t2,j.id).notes['t001'].length===1,'notes belong to one job only');
try{A.addNote(t2,'J1','t001',4,'   ');ok(false,'empty note')}catch(e){ok(/Type a note/.test(e.message),'empty note rejected')}
try{A.addNote(t2,'J1','t001',4,'x'.repeat(2001));ok(false,'long note')}catch(e){ok(/too long/.test(e.message),'over-long note rejected')}
try{A.addNote('bad','J1','t001',4,'hi');ok(false,'auth')}catch(e){ok(e.message==='AUTH','signed-out user cannot add notes')}
ok(A.getJob(t2,'J1').checks['t001'].status==='Done','ticks unaffected by notes');
// quick add from a trade heading = this job only
ok(A.addItems(tj,'J1','job',[{list:4,trade:'Framing',text:'Extra blocking for TV'}])===1&&A.getJob(t2,'J1').items.some(i=>i.text==='Extra blocking for TV'&&i.src==='j'&&i.trade==='Framing'),'add to one trade on this job only');
ok(!A.getJob(t2,j.id).items.some(i=>i.text==='Extra blocking for TV'),'job-only add does not reach other jobs');
