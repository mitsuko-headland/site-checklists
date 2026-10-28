// Backend tests: node tests/backend.test.js  (mocks SpreadsheetApp, Utilities, etc.; prints ok/FAIL per check)
const fs=require('fs');const crypto=require('crypto');
const code=fs.readFileSync(require('path').join(__dirname,'..','src','Code.gs'),'utf8');
function mkSheet(head,rows){const data=[head,...rows];const sh={data,
 getDataRange(){return{getValues:()=>data.map(r=>r.slice())}},getLastColumn(){return Math.max(...data.map(r=>r.length))},getLastRow(){return data.length},getMaxRows(){return 1000},
 getRange(r,c,nr,nc){nr=nr||1;nc=nc||1;return{getValues:()=>{const o=[];for(let i=0;i<nr;i++){const row=data[r-1+i]||[];o.push(Array.from({length:nc},(_,j)=>row[c-1+j]??''));}return o;},
   setValues(v){v.forEach((row,i)=>{data[r-1+i]=data[r-1+i]||[];row.forEach((x,j)=>data[r-1+i][c-1+j]=x);});return this},setValue:v=>{data[r-1]=data[r-1]||[];data[r-1][c-1]=v},setNumberFormat(){},setWrap(){return this},setFontSize(){return this},setFontWeight(){return this}}},
 appendRow(r){data.push(r)},clear(){data.length=0},insertColumnAfter(c){const w=Math.max(...data.map(r=>r.length));data.forEach(r=>{while(r.length<w)r.push('');r.splice(c,0,'');});},deleteRows(s,n){data.splice(s-1,n)},deleteRow(s){data.splice(s-1,1)}};return sh;}
const sheets={
 'Setup':mkSheet(['old'],[['x']]),
 'Users':mkSheet(['Name','Email','Active','Admin'],[['Mark Ellis','m@h.com','Y','Y'],['Joe Site','','Y','N']]),
 'Items':mkSheet(['ItemID','Walk','Trade','Area','Hold','Item','Division','Order','Active'],[['t001',2,'Framing','backframing','HOLD','Backing installed','Div 06-1',1,'Y'],['t002',3,'Drywall','','','Screws mudded','Div 09',2,'Y'],['t003',4,'Paint','','','Paint done','Div 09',3,'Y'],['t004',3,'Energy Consultant','','','Blower door test booked','Div 01',4,'Y']]),
 'Jobs':mkSheet(['JobID','Job Name','Address','Created By','Created At','Active'],[['J1','Old Job','1 St','Mark',new Date(),'Y']]),
 'Job Items':mkSheet(['JobID','ItemID','Type','Walk','Trade','Area','Hold','Item','Added By','Added At','Removed'],[['J1','X1','Item',1,'Plumbing','','HOLD','Job extra','Mark',new Date(),''],['J1','X2','Item',2,'Siding','','','Siding nailed','Mark',new Date(),'']]),
 'Check Log':mkSheet(['JobID','ItemID','Walk','Status','Name','Timestamp'],[['J1','t001',2,'Done','Mark',new Date()]])};
const props={READABLE_V1:'done',STONE_V1:'done',PHASES_V1:'done',GROUNDWORK_REPORT_V1:'done'},cache={};   // readable columns + Stone are tested at the end, so earlier column positions stay put
const ctx={SpreadsheetApp:{getActive:()=>({getSheetByName:n=>sheets[n]||null,insertSheet:n=>{sheets[n]=mkSheet([],[]);sheets[n].data.length=0;sheets[n].setFrozenRows=()=>{};return sheets[n];}})},
 PropertiesService:{getScriptProperties:()=>({getProperty:k=>props[k]||null,setProperty:(k,v)=>props[k]=v})},
 CacheService:{getScriptCache:()=>({get:k=>cache[k]??null,put:(k,v)=>cache[k]=v,remove:k=>delete cache[k]})},
 LockService:{getScriptLock:()=>({waitLock(){if(this.h)throw new Error('deadlock');this.h=1},releaseLock(){this.h=0}})},HtmlService:{},
 Utilities:{getUuid:()=>crypto.randomUUID(),computeHmacSha256Signature:(p,s)=>[...crypto.createHmac('sha256',s).update(p).digest()],
  base64EncodeWebSafe:x=>Buffer.from(typeof x==='string'?x:Buffer.from(x.map(b=>b&255))).toString('base64').replace(/\+/g,'-').replace(/\//g,'_'),
  base64DecodeWebSafe:s=>[...Buffer.from(s.replace(/-/g,'+').replace(/_/g,'/'),'base64')],newBlob:b=>({getDataAsString:()=>Buffer.from(b.map(x=>x&255)).toString()}),formatDate:d=>d.toISOString().slice(5,16)}};
const A=new Function(...Object.keys(ctx),code+';return {editItem,setup,login,whoAmI,getJobs,createJob,getJob,setCheck,addItems,removeItem,addNote,startup,cleanupReport_,cleanupRun_,setListHidden,restoreHidden_,readableUpgrade_,stoneUpgrade_,phasesUpgrade_,groundworkReport_};')(...Object.values(ctx));
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
console.log('   items',d.items.map(i=>i.id+':'+i.list).join(' '),'check',JSON.stringify(d.checks));ok(d.items.length===10&&d.checks['t001'].status==='Done','old job keeps its tick after upgrade');
ok(d.items.find(i=>i.id==='t001').critical&&d.items.find(i=>i.id==='X1').critical,'critical flags read');
const j=A.createJob(t,'Chen','12 Main');ok(A.getJob(t,j.id).items.filter(i=>i.list===1).length===4,'new job gets startup checklist');
ok(A.setCheck(t,j.id,'S001',1,'Done').name==='Mark Ellis','tick on checklist 1');
ok(A.addItems(t,j.id,'all',[{list:2,trade:'Demo',critical:true,text:'Hazmat report on site'}])===1,'admin adds to Demolition (inside lock, no deadlock)');
ok(A.getJob(t,j.id).items.find(i=>i.list===2).critical,'new item critical');
A.removeItem(t,j.id,'t002','m',5);ok(!A.getJob(t,j.id).items.some(i=>i.id==='t002'),'hide master item for job');
U[1][2]='N';try{A.getJobs(t);ok(false,'blk')}catch(e){ok(e.message==='AUTH','blocked user loses access')}

U[1][2]='Y';
function r3t(){U[2][4]='5555';delete cache['fail:joe site'];return A.login('Joe Site','5555').token;}
const t2=A.login('Mark Ellis','0123').token;
// editing a master-list task changes it on THIS job only (anyone can edit); Items tab untouched
const Jother=A.createJob(t2,'Other job','');
let n=A.editItem(r3t(),'J1','t003','m',{text:'Paint touch-ups complete',trade:'Painting',area:'ceilings',critical:true});
ok(n===4,'non-admin edit of a master task changes 4 fields');
const it=sheets['Items'].data.find(r=>r[0]==='t003');ok(it[5]==='Paint done'&&it[2]==='Paint','master list on the Items tab is unchanged');
{const e=A.getJob(t2,'J1').items.find(i=>i.id==='t003'),o=A.getJob(t2,Jother.id).items.find(i=>i.id==='t003');
 ok(e.text==='Paint touch-ups complete'&&e.trade==='Painting'&&e.area==='ceilings'&&e.critical&&e.edited,'edited job shows the new wording, marked edited');
 ok(o.text==='Paint done'&&!o.edited,'other jobs keep the master wording');}
ok(sheets['Job Items'].data.filter(r=>r[2]==='EDIT'&&r[1]==='t003'&&r[0]==='J1').length===1,'one EDIT row for that job');
const L=sheets['Edit Log'].data;ok(L[0][0]==='Timestamp'&&L.filter(r=>r[1]==='Joe Site'&&r[2]==='J1'&&r[3]==='t003').length===4&&L.some(r=>r[4]==='Wording'&&r[5]==='Paint done'&&r[6]==='Paint touch-ups complete'&&r[1]==='Joe Site'),'Edit Log records job, old -> new and who');
ok(A.editItem(t2,'J1','t003','m',{text:'Paint touch-ups complete',trade:'Painting',area:'ceilings',critical:true})===0,'no-op edit logs nothing');
ok(A.editItem(t2,'J1','t003','m',{text:'Paint touch-ups done',trade:'Painting',area:'ceilings',critical:true})===1&&sheets['Job Items'].data.filter(r=>r[2]==='EDIT'&&r[1]==='t003').length===1&&A.getJob(t2,'J1').items.find(i=>i.id==='t003').text==='Paint touch-ups done','second edit updates the same EDIT row');
ok(!A.getJob(t2,'J1').items.some(i=>i.src==='j'&&i.id==='t003'),'EDIT rows never show as extra tasks');
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
// job-only tasks go at the top of their trade: newest batch first, typed order within a batch
{const J=A.createJob(t2,'Order test','');
 A.addItems(t2,J.id,'job',[{list:4,trade:'Framing',text:'First batch A'},{list:4,trade:'Framing',text:'First batch B'}]);
 const jr=sheets['Job Items'].data;jr.filter(r=>r[0]===J.id).forEach(r=>r[9]=new Date(Date.now()-60000));   // first batch a minute ago
 A.addItems(t2,J.id,'job',[{list:4,trade:'Framing',text:'Second batch A'},{list:4,trade:'Framing',text:'Second batch B'}]);
 const fr=A.getJob(t2,J.id).items.filter(i=>i.list===4&&i.trade==='Framing').sort((a,b)=>a.order-b.order).map(i=>i.text);
 ok(fr.slice(0,4).join('|')==='Second batch A|Second batch B|First batch A|First batch B','newest added tasks first, typed order kept within a batch');
 ok(fr.length>4&&A.getJob(t2,J.id).items.filter(i=>i.src==='j').every(i=>i.order<0),'added tasks sort ahead of master-list tasks');}
// start-up in one request: who is signed in + jobs + the open job
{const t=A.login('Mark Ellis',sheets['Users'].data[1][4]).token;
 const s=A.startup(t,'J1');
 ok(s.me.name==='Mark Ellis'&&s.me.admin===true,'startup returns who is signed in');
 ok(s.jobs.length===A.getJobs(t).length&&s.jobs.some(j=>j.id==='J1'),'startup returns the job list');
 ok(s.job&&s.job.job.id==='J1'&&JSON.stringify(s.job)===JSON.stringify(A.getJob(t,'J1')),'startup returns the same job data as getJob');
 ok(A.startup(t,'').job===null&&A.startup(t,'NOPE').job===null,'startup without a job, or with a missing job, returns no job');
 try{A.startup('bad','J1');ok(false,'auth')}catch(e){ok(e.message==='AUTH','startup refuses a bad sign-in')}
 // ticks no longer take the lock, but still log one row each and the latest wins
 const before=sheets['Check Log'].data.length;
 A.setCheck(t,'J1','t002',5,'Done');A.setCheck(t,'J1','t002',5,'Cleared');A.setCheck(t,'J1','t002',5,'Done');
 ok(sheets['Check Log'].data.length===before+3,'each tick adds one Check Log row');
 ok(A.getJob(t,'J1').checks['t002'].status==='Done','latest tick wins');
 ok(sheets['Check Log'].data.slice(-1)[0][4]==='Mark Ellis','tick logged with the signed-in name');}
// trade codes: Trades tab, one-time moves, codes sent to the app
{const T=sheets['Trades'];
 ok(T&&T.data[0].join()==='Code,Trade'&&T.data.length===31,'Trades tab created with 30 trades');
 ok(T.data.some(r=>r[0]==='1050'&&r[1]==='Site Supervisor')&&T.data.some(r=>r[0]==='8202'&&r[1]==='Doors'),'Trades tab has the agreed codes');
 const I=sheets['Items'].data,JI=sheets['Job Items'].data;
 ok(I.find(r=>r[0]==='t004')[2]==='Site Supervisor','Energy Consultant item moved to Site Supervisor');
 ok(JI.find(r=>r[1]==='X2')[4]==='Exterior Carpentry','job-only Siding item moved to Exterior Carpentry');
 ok(['S001','S002','S003','S004'].every(id=>I.find(r=>r[0]===id)[2]==='Mobilize & Startup Site Protection'),'4 startup items moved to Mobilize & Startup Site Protection');
 const EL=sheets['Edit Log'].data.filter(r=>r[1]==='Setup (trade codes)');
 ok(EL.length===6&&EL.some(r=>r[3]==='t004'&&r[5]==='Energy Consultant'&&r[6]==='Site Supervisor')&&EL.some(r=>r[2]==='J1'&&r[3]==='X2'),'every move is in the Edit Log (master and job-only)');
 const rep=JSON.parse(props['TRADES_V1_REPORT']);
 ok(rep.moved['Energy Consultant → Site Supervisor']===1&&rep.moved['Site Supervisor → Mobilize & Startup Site Protection']===4,'report counts moves');
 ok(rep.uncoded.includes('Paint')&&!rep.uncoded.includes('Framing'),'report lists trades with no code');
 ok(sheets['Setup'].data.some(r=>String(r[0]).includes('TRADE CODES UPDATE')),'Setup tab shows the trade codes report');
 delete cache['schema'];A.setup();
 ok(sheets['Edit Log'].data.filter(r=>r[1]==='Setup (trade codes)').length===6&&sheets['Trades'].data.length===31,'trade moves run only once');
 const t=A.login('Mark Ellis',sheets['Users'].data[1][4]).token;const d=A.getJob(t,'J1');
 ok(d.trades.find(x=>x.name==='Framing').code==='6101'&&d.trades.length===30,'job data includes the trade codes');}
// job-only tasks say who added them and the day; master-list tasks don't
{const t=A.login('Mark Ellis',sheets['Users'].data[1][4]).token;const J=A.createJob(t,'Added-by test','');
 A.addItems(t,J.id,'job',[{list:3,trade:'Framing',text:'Extra king stud'}]);
 const d=A.getJob(t,J.id),x=d.items.find(i=>i.text==='Extra king stud'),m=d.items.find(i=>i.src==='m');
 ok(x.addedBy==='Mark Ellis'&&/^\d\d-\d\d/.test(x.addedOn),'job-only task carries who added it and the day');
 ok(!('addedBy' in m)&&!('addedOn' in m),'master-list tasks carry no added-by');}
// anyone can remove a task from one job; the removal is logged
{const t=r3t();const J=A.createJob(t,'Remove test','');
 A.removeItem(t,J.id,'t002','m',5);
 ok(!A.getJob(t,J.id).items.some(i=>i.id==='t002')&&A.getJob(t,'J1').items.some(i=>i.id==='t002'),'non-admin removes a master task from one job only');
 ok(sheets['Edit Log'].data.some(r=>r[1]==='Joe Site'&&r[2]===J.id&&r[3]==='t002'&&r[4]==='Removed from job'&&r[5]==='Screws mudded'),'removal is in the Edit Log with who and the task wording');}
// Park clean-up summary: lists only one person's changes on one job, changes nothing
{const U=sheets['Users'].data;U.push(['Mitsuko','','Y','Y','1111']);U.push(['Sam','','Y','N','2222']);
 const tm=A.login('Mitsuko','1111').token,ts=A.login('Sam','2222').token;
 const P=A.createJob(tm,'Park','12 Park St');A.createJob(tm,'Parkview','');
 A.setCheck(ts,P.id,'t001',4,'Done');           // Sam ticks
 A.setCheck(tm,P.id,'t001',4,'Cleared');        // Mitsuko unticks it
 A.setCheck(tm,P.id,'t002',5,'Done');           // Mitsuko ticks
 A.addNote(tm,P.id,'t002',5,'test note');A.addNote(ts,P.id,'t002',5,'real note from Sam');
 A.addItems(tm,P.id,'job',[{list:3,trade:'Framing',text:'Mitsuko test task'}]);
 A.addItems(ts,P.id,'job',[{list:3,trade:'Framing',text:'Sam real task'}]);
 const mt=A.getJob(tm,P.id).items.find(i=>i.text==='Mitsuko test task');A.setCheck(ts,P.id,mt.id,3,'Done');   // Sam ticks Mitsuko's task
 A.removeItem(tm,P.id,'t003','m',6);
 A.editItem(tm,P.id,'S001','m',{text:'Site signage (Park)',trade:'Mobilize & Startup Site Protection',area:'',critical:false});
 sheets['Edit Log'].data.push([new Date(),'Mitsuko','(all jobs)','t002','Wording','Old wording','Screws mudded']);   // an older master edit
 const before=JSON.stringify(Object.keys(sheets).map(k=>[k,sheets[k].data.length]));
 A.cleanupReport_(ctx.SpreadsheetApp.getActive(),'Park','Mitsuko','Park Cleanup');
 const R=sheets['Park Cleanup'].data.map(r=>r.join(' | ')),txt=R.join('\n');
 ok(R[1].includes('Park ('+P.id+')')&&!R[1].includes('Parkview'),'exact job name "Park" chosen over "Parkview"');
 ok(txt.includes('1. YOUR TICKS AND UNTICKS (will be deleted from Check Log) – 2'),'lists 2 of Mitsuko\'s ticks, none of Sam\'s');
 ok(R.some(r=>r.includes('Backing installed')&&r.includes('Cleared')&&r.includes('Afterwards: Done by Sam')),'shows what a task goes back to (Sam\'s tick)');
 ok(R.some(r=>r.includes('Screws mudded')&&r.includes('Afterwards: not done')),'shows a task going back to not done');
 ok(txt.includes('2. YOUR NOTES (will be deleted) – 1')&&txt.includes('test note')&&!txt.includes('real note from Sam'),'lists only her notes');
 ok(txt.includes('3. TASKS YOU ADDED TO THIS JOB (will be removed) – 1')&&R.some(r=>r.includes('Mitsuko test task')&&r.includes('CHECK: Sam ticked this task'))&&!txt.includes('Sam real task'),'lists her added task and flags Sam\'s tick on it');
 ok(txt.includes('4. TASKS YOU REMOVED FROM THIS JOB (will come back) – 1'),'lists her removal');
 ok(txt.includes('5. YOUR EDITS ON THIS JOB (will be undone) – 1')&&txt.includes('"Site signage" → "Site signage (Park)"'),'lists her job edit with old and new');
 ok(txt.includes('6. YOUR MASTER-LIST EDITS')&&txt.includes('"Old wording" → "Screws mudded"'),'lists her older master-list edits');
 ok(/7\. MASTER-LIST TASKS ADDED THROUGH THE APP[^\n]*– 1/.test(txt)&&txt.includes('Hazmat report on site'),'lists master tasks added through the app');
 ok(JSON.stringify(Object.keys(sheets).filter(k=>k!=='Park Cleanup').map(k=>[k,sheets[k].data.length]))===JSON.stringify(JSON.parse(before).filter(x=>x[0]!=='Park Cleanup')),'summary changes nothing else');
 // a job-only task by Sam that Mitsuko edited and then removed (before the clearing)
 const st=A.getJob(ts,P.id).items.find(i=>i.text==='Sam real task');
 A.editItem(tm,P.id,st.id,'j',{text:'Sam real task (Mitsuko test edit)',trade:'Framing',area:'',critical:false});A.removeItem(tm,P.id,st.id,'j',3);
 const other=A.createJob(tm,'Other site','');A.setCheck(tm,other.id,'t002',5,'Done');A.addNote(tm,other.id,'t002',5,'keep me');
 const res=A.cleanupRun_(ctx.SpreadsheetApp.getActive(),'Park','Mitsuko','Park Cleanup');
 const dp=A.getJob(ts,P.id);
 ok(!sheets['Check Log'].data.some(r=>r[0]===P.id&&r[4]==='Mitsuko')&&sheets['Check Log'].data.some(r=>r[0]===P.id&&r[4]==='Sam'),'her Park ticks deleted, Sam\'s kept');
 ok(dp.checks['t001']&&dp.checks['t001'].status==='Done'&&dp.checks['t001'].name==='Sam'&&!dp.checks['t002'],'tasks go back to Sam\'s tick, or not done');
 ok(!sheets['Notes'].data.some(r=>r[0]===P.id&&r[4]==='Mitsuko')&&sheets['Notes'].data.some(r=>r[3]==='real note from Sam'),'her notes deleted, Sam\'s kept');
 ok(dp.items.some(i=>i.text==='Mitsuko test task'),'her task that Sam ticked is kept');
 ok(dp.items.some(i=>i.id==='t003'),'task she removed is back');
 ok(dp.items.find(i=>i.id==='S001').text==='Site signage'&&!dp.items.find(i=>i.id==='S001').edited,'her edit to a master task on Park is undone');
 ok(dp.items.some(i=>i.text==='Sam real task'),'Sam\'s task she edited and removed is back with its original wording');
 ok(!sheets['Edit Log'].data.some(r=>r[2]===P.id&&r[1]==='Mitsuko')&&sheets['Edit Log'].data.some(r=>r[2]==='(all jobs)'&&r[1]==='Mitsuko'),'her Park Edit Log rows deleted; master-list edit rows untouched');
 ok(sheets['Check Log'].data.some(r=>r[0]===other.id&&r[4]==='Mitsuko')&&sheets['Notes'].data.some(r=>r[3]==='keep me'),'her work on other jobs untouched');
 ok(sheets['Items'].data.some(r=>r[5]==='Hazmat report on site'),'master list untouched');
 const B=sheets['Park Cleanup – Removed rows'].data;ok(B.length>1&&B.some(r=>r[0]==='Check Log')&&B.some(r=>r[0]==='Notes'&&r.includes('test note')),'deleted rows copied to the backup tab');
 const R2=sheets['Park Cleanup'].data.map(r=>r.join(' | ')).join('\n');
 ok(R2.startsWith('PARK CLEAN-UP – DONE')&&R2.includes('Kept your task "Mitsuko test task" – Sam ticked or noted it.'),'summary tab rewritten with the result and what was kept');
 ok(res.ticks===2&&res.notes===1&&res.unhidden===1&&res.unedited===1&&res.restored===1&&res.reverted===1,'counts: 2 ticks, 1 note, 1 back, 1 master edit undone, 1 task restored, 1 edit reversed');}
// admins hide a whole checklist on one job; logged; nothing deleted
{const ta=A.login('Mark Ellis',sheets['Users'].data[1][4]).token,tj=r3t();
 const J1=A.createJob(ta,'Hide test',''),J2=A.createJob(ta,'Other hide test','');
 A.setCheck(tj,J1.id,'t002',5,'Done');
 try{A.setListHidden(tj,J1.id,2,true);ok(false,'non-admin hide')}catch(e){ok(/Only admins/.test(e.message),'non-admin cannot hide a checklist')}
 ok(A.setListHidden(ta,J1.id,5,true)===true&&JSON.stringify(A.getJob(ta,J1.id).hiddenLists)==='[5]','admin hides checklist 5 on one job');
 ok(JSON.stringify(A.getJob(ta,J2.id).hiddenLists)==='[]','other jobs unaffected');
 ok(A.getJob(ta,J1.id).checks['t002'].status==='Done'&&A.getJob(ta,J1.id).items.some(i=>i.list===5),'ticks and tasks kept while hidden');
 const n0=sheets['Edit Log'].data.length;A.setListHidden(ta,J1.id,5,true);ok(sheets['Edit Log'].data.length===n0,'hiding twice logs nothing new');
 ok(sheets['Edit Log'].data.some(r=>r[1]==='Mark Ellis'&&r[2]===J1.id&&r[3]==='Checklist 5'&&r[6]==='Hidden'),'hide is in the Edit Log');
 A.setListHidden(ta,J1.id,5,false);
 ok(JSON.stringify(A.getJob(ta,J1.id).hiddenLists)==='[]'&&sheets['Edit Log'].data.some(r=>r[2]===J1.id&&r[3]==='Checklist 5'&&r[6]==='Shown'),'unhide shows it again and is logged');
 ok(!A.getJob(ta,J1.id).items.some(i=>i.src==='j'&&/^L\d$/.test(i.id)),'hide rows never show as tasks');}
// one-time restore of a task removed by mistake: only that HIDE row goes
{const ta=A.login('Mark Ellis',sheets['Users'].data[1][4]).token,tmk=A.login('Mitsuko','1111').token;
 const J=A.createJob(ta,'Restore test',''),K=A.createJob(ta,'Keep hidden','');
 A.removeItem(tmk,J.id,'t002','m',5);A.removeItem(tmk,K.id,'t002','m',5);
 ok(!A.getJob(ta,J.id).items.some(i=>i.id==='t002'),'task removed before the fix');
 ok(A.restoreHidden_(ctx.SpreadsheetApp.getActive(),J.id,'t002','Mitsuko')===1,'one HIDE row deleted');
 ok(A.getJob(ta,J.id).items.some(i=>i.id==='t002')&&!A.getJob(ta,K.id).items.some(i=>i.id==='t002'),'task back on that job only');
 ok(sheets['Edit Log'].data.some(r=>r[2]===J.id&&r[3]==='t002'&&r[4]==='Restored to job'),'restore is logged');}
// readable Job Name + Task columns on the log tabs, then filled on every new row
{const ss=ctx.SpreadsheetApp.getActive(),ta=A.login('Mark Ellis',sheets['Users'].data[1][4]).token;
 const P=sheets['Jobs'].data.find(r=>r[1]==='Park')[0];
 A.readableUpgrade_(ss);
 const hd=n=>sheets[n].data[0];
 ok(['Job Items','Check Log','Notes','Edit Log'].every(n=>{const h=hd(n);return h[h.indexOf('JobID')+1]==='Job Name'&&h[h.indexOf('ItemID')+1]==='Task';}),'Job Name after JobID and Task after ItemID on all four log tabs');
 const CL=sheets['Check Log'].data,ci=c=>hd('Check Log').indexOf(c);
 ok(CL.slice(1).every(r=>r[ci('Job Name')]===(sheets['Jobs'].data.find(j=>j[0]===r[ci('JobID')])||[])[1]),'existing Check Log rows filled with the job name');
 ok(CL.slice(1).some(r=>r[ci('ItemID')]==='t001'&&r[ci('Task')]==='Backing installed'),'existing rows filled with the task wording');
 const EL=sheets['Edit Log'].data,ei=c=>hd('Edit Log').indexOf(c);
 ok(EL.slice(1).some(r=>r[ei('JobID')]==='(all jobs)'&&r[ei('Job Name')]==='All jobs')&&EL.slice(1).some(r=>/^Checklist 5/.test(r[ei('ItemID')])&&r[ei('Task')]==='Checklist 5 – Pre-Drywall Review: Framing & Trades Rough In'),'"(all jobs)" and checklist rows get readable labels');
 A.readableUpgrade_(ss);ok(hd('Check Log').filter(c=>c==='Job Name').length===1,'running it again adds no extra columns');
 const J=A.createJob(ta,'Butt','');A.setCheck(ta,J.id,'t001',4,'Done');
 const last=sheets['Check Log'].data.slice(-1)[0];ok(last[ci('Job Name')]==='Butt'&&last[ci('Task')]==='Backing installed'&&last[ci('Status')]==='Done','new ticks are written with job name and task');
 A.addItems(ta,J.id,'job',[{list:3,trade:'Framing',text:'Readable task'}]);const ji=sheets['Job Items'].data.slice(-1)[0],jx=c=>hd('Job Items').indexOf(c);
 ok(ji[jx('Job Name')]==='Butt'&&ji[jx('Task')]==='Readable task','new job tasks carry job name and task');
 A.editItem(ta,J.id,'t001','m',{text:'Backing installed (Butt)',trade:'Framing',area:'',critical:false});
 A.setCheck(ta,J.id,'t001',4,'Cleared');ok(sheets['Check Log'].data.slice(-1)[0][ci('Task')]==='Backing installed (Butt)','task label uses this job\'s own wording');
 ok(A.getJob(ta,J.id).checks['t001']===null&&A.getJob(ta,J.id).items.some(i=>i.text==='Readable task'),'app still reads everything after the new columns');
 // Stone
 sheets['Items'].data.push(['t900',6,'Stone Countertops','','','Countertop seams checked','',99,'Y']);
 sheets['Trades'].data.find(r=>r[1]==='Stone')[1]='Stone Countertops';   // as on the live sheet today
 A.stoneUpgrade_(ss);
 const T=sheets['Trades'].data;ok(T.some(r=>r[0]==='12370'&&r[1]==='Stone')&&!T.some(r=>r[1]==='Stone Countertops'),'Trades tab: 12370 Stone');
 const ix=sheets['Items'].data[0].indexOf('Trade');ok(sheets['Items'].data.find(r=>r[0]==='t900')[ix]==='Stone','Stone Countertops tasks moved to Stone');
 ok(A.getJob(ta,J.id).trades.find(x=>x.name==='Stone').code==='12370','heading code for Stone is 12370');}
// checklist 0 – Project management to do – admins only
{const ta=A.login('Mark Ellis',sheets['Users'].data[1][4]).token,tc=r3t();const J=A.createJob(ta,'PM test','');
 ok(A.addItems(ta,J.id,'job',[{list:0,trade:'Site Supervisor',text:'Order windows'}])===1,'admin adds a task to checklist 0');
 try{A.addItems(tc,J.id,'job',[{list:0,trade:'Site Supervisor',text:'Sneaky'}]);ok(false,'crew add 0')}catch(e){ok(/Only admins/.test(e.message),'crew cannot add to checklist 0')}
 const t0=A.getJob(ta,J.id).items.find(i=>i.text==='Order windows');ok(t0&&t0.list===0,'admin sees the checklist 0 task');
 A.setCheck(ta,J.id,t0.id,0,'Done');A.addNote(ta,J.id,t0.id,0,'PO 123');
 const dc=A.getJob(tc,J.id);ok(!dc.items.some(i=>i.list===0)&&!dc.checks[t0.id]&&!dc.notes[t0.id],'crew get no checklist 0 tasks, ticks or notes');
 ok(!A.startup(tc,J.id).job.items.some(i=>i.list===0)&&A.startup(ta,J.id).job.items.some(i=>i.list===0),'same at start-up');
 ok(sheets['Notes'].data.slice(-1)[0][sheets['Notes'].data[0].indexOf('Checklist')]===0,'note on checklist 0 keeps checklist number 0');
 sheets['Items'].data.push(['t950','','Framing','','','No checklist number','',999,'Y']);
 ok(!A.getJob(ta,J.id).items.some(i=>i.id==='t950'&&i.list===0),'a blank Checklist cell never lands in checklist 0');}

// move a task to another checklist (Edit form, this job only)
{const ta=A.login('Mark Ellis',sheets['Users'].data[1][4]).token,tc=r3t();const J=A.createJob(ta,'Move test',''),K=A.createJob(ta,'Other move job','');
 A.addItems(ta,J.id,'job',[{list:0,trade:'Site Supervisor',text:'Wrong list task'},{list:3,trade:'Framing',text:'Crew job task'}]);
 let d=A.getJob(ta,J.id);const w=d.items.find(i=>i.text==='Wrong list task'),c=d.items.find(i=>i.text==='Crew job task');
 A.setCheck(ta,J.id,w.id,0,'Done');A.addNote(ta,J.id,w.id,0,'keep me');
 ok(A.editItem(ta,J.id,w.id,'j',{text:'Wrong list task',trade:'Site Supervisor',area:'',critical:false,list:2})===1,'admin moves a job-only task out of checklist 0');
 d=A.getJob(ta,J.id);const w2=d.items.find(i=>i.id===w.id);ok(w2.list===2&&d.checks[w.id].status==='Done'&&d.notes[w.id].length===1,'moved task keeps its tick and notes');
 ok(A.getJob(tc,J.id).items.some(i=>i.id===w.id&&i.list===2)&&A.getJob(tc,J.id).checks[w.id],'crew now see it in checklist 2');
 const EL=sheets['Edit Log'].data,eh=EL[0],e=EL.slice(-1)[0];ok(e[eh.indexOf('Field')]==='Checklist'&&e[eh.indexOf('Old Value')]==='0 Project management to do'&&e[eh.indexOf('New Value')]==='2 Demolition'&&e[eh.indexOf('Name')]==='Mark Ellis','Edit Log: Checklist 0 → 2 with names and who');
 ok(A.editItem(tc,J.id,c.id,'j',{text:'Crew job task',trade:'Framing',area:'',critical:false,list:4})===1&&A.getJob(tc,J.id).items.find(i=>i.id===c.id).list===4,'crew can move a job-only task between checklists');
 ok(A.editItem(tc,J.id,c.id,'j',{text:'Crew job task',trade:'Framing',area:'',critical:false,list:4})===0,'same checklist = no change');
 ok(A.editItem(tc,J.id,c.id,'j',{text:'Crew job task',trade:'Framing',area:''})===0,'no checklist sent = stays put');
 try{A.editItem(tc,J.id,c.id,'j',{text:'Crew job task',trade:'Framing',list:0});ok(false,'crew into 0')}catch(x){ok(/Only admins/.test(x.message),'crew cannot move a task into checklist 0')}
 try{A.editItem(tc,J.id,c.id,'j',{text:'Crew job task',trade:'Framing',list:9});ok(false,'bad list')}catch(x){ok(/Unknown checklist/.test(x.message),'unknown checklist rejected')}
 A.setListHidden(ta,J.id,5,true);
 try{A.editItem(tc,J.id,c.id,'j',{text:'Crew job task',trade:'Framing',list:5});ok(false,'crew into hidden')}catch(x){ok(/hidden/.test(x.message),'crew cannot move into a checklist hidden on this job')}
 ok(A.editItem(ta,J.id,c.id,'j',{text:'Crew job task',trade:'Framing',list:5})===1,'admin can move into a hidden checklist');A.setListHidden(ta,J.id,5,false);
 // master task: this job only
 const before=sheets['Items'].data.find(r=>r[0]==='t001').slice();const ml=A.getJob(ta,J.id).items.find(i=>i.id==='t001');
 ok(A.editItem(tc,J.id,'t001','m',{text:ml.text,trade:ml.trade,area:ml.area,critical:ml.critical,list:6})===1,'crew move a master task on one job');
 ok(A.getJob(tc,J.id).items.find(i=>i.id==='t001').list===6&&A.getJob(tc,K.id).items.find(i=>i.id==='t001').list===ml.list,'moved on this job only; other jobs unchanged');
 ok(JSON.stringify(sheets['Items'].data.find(r=>r[0]==='t001'))===JSON.stringify(before),'master list on Items tab unchanged');
 ok(A.editItem(tc,J.id,'t001','m',{text:ml.text+' (edited)',trade:ml.trade,area:ml.area,critical:ml.critical})===1&&A.getJob(tc,J.id).items.find(i=>i.id==='t001').list===6,'later wording edit keeps the moved checklist');
 ok(A.editItem(tc,J.id,'t001','m',{text:ml.text+' (edited)',trade:ml.trade,area:ml.area,critical:ml.critical,list:ml.list})===1&&A.getJob(tc,J.id).items.find(i=>i.id==='t001').list===ml.list,'moved back again');
 {const H=sheets['Job Items'].data[0],x=c=>H.indexOf(c);ok(sheets['Job Items'].data.filter(r=>r[x('Type')]==='EDIT'&&r[x('ItemID')]==='t001'&&r[x('JobID')]===J.id).length===1,'still one EDIT row per job + task');}
 // master task in checklist 0 moved by admin, hidden from crew before
 A.editItem(ta,J.id,'t001','m',{text:ml.text,trade:ml.trade,list:0});ok(!A.getJob(tc,J.id).items.some(i=>i.id==='t001')&&A.getJob(ta,J.id).items.find(i=>i.id==='t001').list===0,'master task moved into 0 is hidden from crew on that job');
 try{A.editItem(tc,J.id,'t001','m',{text:ml.text,trade:ml.trade,list:1});ok(false,'crew out of 0')}catch(x){ok(/Only admins/.test(x.message),'crew cannot move a task out of checklist 0')}}
// "Change on all jobs" (admins, master-list tasks): Items tab updated; every job's own version replaced and logged
{const ta=A.login('Mark Ellis',sheets['Users'].data[1][4]).token,tc=r3t();const J=A.createJob(ta,'Global A',''),K=A.createJob(ta,'Global B',''),L=A.createJob(ta,'Global C','');
 const IH=sheets['Items'].data[0],ix=c=>IH.indexOf(c),row=()=>sheets['Items'].data.find(r=>r[0]==='t002');const m0=A.getJob(ta,J.id).items.find(i=>i.id==='t002');
 A.editItem(tc,K.id,'t002','m',{text:'B own wording',trade:m0.trade,area:'kitchen',critical:m0.critical});   // job B has its own version
 ok(A.getJob(ta,K.id).items.find(i=>i.id==='t002').edited,'job B has its own version');
 try{A.editItem(tc,J.id,'t002','m',{text:'Crew global',trade:m0.trade,global:true});ok(false,'crew global')}catch(e){ok(/Only admins/.test(e.message),'crew cannot change a task on all jobs')}
 try{const x=A.getJob(ta,J.id).items.find(i=>i.src==='j')||(A.addItems(ta,J.id,'job',[{list:2,trade:'Framing',text:'J only'}]),A.getJob(ta,J.id).items.find(i=>i.text==='J only'));A.editItem(ta,J.id,x.id,'j',{text:'x',trade:'Framing',global:true});ok(false,'job-only global')}catch(e){ok(/master-list/.test(e.message),'job-only tasks cannot be changed on all jobs')}
 const n=A.editItem(ta,J.id,'t002','m',{text:'New master wording',trade:m0.trade,area:'',critical:true,list:3,global:true});
 ok(n>=3,'admin changes the task on all jobs');
 ok(row()[ix('Item')]==='New master wording'&&row()[ix('Critical')]==='CRITICAL'&&Number(row()[ix('Checklist')])===3,'Items tab updated (wording, critical, checklist)');
 [J,K,L].forEach(j=>{const it=A.getJob(ta,j.id).items.find(i=>i.id==='t002');ok(it.text==='New master wording'&&it.list===3&&!it.edited,'every job shows the new version: '+j.name);});
 const EL=sheets['Edit Log'].data,eh=c=>EL[0].indexOf(c);
 ok(EL.some(r=>r[eh('JobID')]===K.id&&r[eh('Field')]==='Job version replaced (all-jobs edit)'&&/B own wording/.test(r[eh('Old Value')])&&/kitchen/.test(r[eh('Old Value')])),'job B\'s replaced version is kept in Edit Log');
 ok(EL.some(r=>r[eh('JobID')]==='(all jobs)'&&r[eh('Field')]==='Wording'&&r[eh('New Value')]==='New master wording'&&r[eh('Name')]==='Mark Ellis')&&EL.some(r=>r[eh('JobID')]==='(all jobs)'&&r[eh('Field')]==='Checklist'&&r[eh('New Value')]==='3 Groundwork & Slab Prep'),'all-jobs changes logged with who');
 ok(A.editItem(ta,J.id,'t002','m',{text:'New master wording',trade:m0.trade,area:'',critical:true,list:3,global:true})===0,'same values again = no change');
 ok(A.editItem(ta,J.id,'t002','m',{text:'Local only',trade:m0.trade,area:'',critical:true})===1&&row()[ix('Item')]==='New master wording'&&A.getJob(ta,K.id).items.find(i=>i.id==='t002').text==='New master wording','box not ticked = this job only, master unchanged');}

// 2026-10-28: new checklist 3 Groundwork & Slab Prep – old 3–6 move to 4–7 on every tab; Edit Log untouched
{const T={
  'Items':mkSheet(['ItemID','Checklist','Trade','Item'],[['a',0,'X','zero'],['b',2,'X','demo'],['c',3,'X','pre'],['d',4,'X','rough'],['e',5,'X','dry'],['f',6,'X','fin'],['g','','X','blank']]),
  'Job Items':mkSheet(['JobID','Job Name','ItemID','Task','Type','Checklist','Item','Removed'],[['J','Jn','x1','t','Item',3,'own',''],['J','Jn','L4','Checklist 4 – Rough In & Pre-Drywall','LIST',4,'',''],['J','Jn','L2','Checklist 2 – Demolition','LIST',2,'',''],['J','Jn','d','rough','EDIT',6,'moved','']]),
  'Check Log':mkSheet(['JobID','ItemID','Checklist','Status'],[['J','c',3,'Done'],['J','b',2,'Done'],['J','f',6,'Done']]),
  'Notes':mkSheet(['JobID','ItemID','Checklist','Note'],[['J','e',5,'n'],['J','a',0,'n0']])};
 const ss={getSheetByName:n=>T[n]||null};
 const elBefore=sheets['Edit Log'].data.length;
 const cnt=A.phasesUpgrade_(ss);
 ok(T['Items'].data.slice(1).map(r=>r[1]).join()==='0,2,4,5,6,7,','Items: 0–2 and blank stay, 3–6 become 4–7');
 ok(T['Job Items'].data[1][5]===4&&T['Job Items'].data[4][5]===7,'Job Items: job-only task and EDIT checklist moved');
 ok(T['Job Items'].data[2][2]==='L5'&&T['Job Items'].data[2][5]===5&&T['Job Items'].data[2][3]==='Checklist 5 – Pre-Drywall Review: Framing & Trades Rough In','hidden checklist row moves with it (L4 → L5, label updated)');
 ok(T['Job Items'].data[3][2]==='L2'&&T['Job Items'].data[3][5]===2,'hidden Demolition row unchanged');
 ok(T['Check Log'].data.slice(1).map(r=>r[2]).join()==='4,2,7'&&T['Notes'].data.slice(1).map(r=>r[2]).join()==='6,0','Check Log and Notes renumbered');
 ok(cnt['Items']===4&&cnt['Job Items']===3&&cnt['Check Log']===2&&cnt['Notes']===1,'counts reported');
 const EL=sheets['Edit Log'].data,eh=c=>EL[0].indexOf(c);ok(EL.length===elBefore+1&&EL[EL.length-1][eh('Field')]==='Checklists renumbered','one Edit Log line records the renumber');
 // real sheet: the upgrade runs once only
 delete props['PHASES_V1'];delete cache['schema'];const IC=sheets['Items'].data[0].indexOf('Checklist'),snap=sheets['Items'].data.map(r=>r[IC]).join();A.setup();
 const after=sheets['Items'].data.map(r=>r[IC]).join();ok(props['PHASES_V1']==='done'&&after!==snap,'upgrade runs from setup and is marked done');
 delete cache['schema'];A.setup();ok(sheets['Items'].data.map(r=>r[IC]).join()===after,'second run changes nothing');}
// read-only groundwork list
{const I=sheets['Items'].data,h=I[0],row=h.map(c=>({ItemID:'g1',Checklist:4,Trade:'Concrete',Item:'Gravel compacted under slab',Active:'Y'})[c]??'');I.push(row);
 const row2=h.map(c=>({ItemID:'g2',Checklist:6,Trade:'Drywall',Item:'Tape and mud',Active:'Y'})[c]??'');I.push(row2);
 const before=JSON.stringify(I);const n=A.groundworkReport_(SpreadsheetApp_());
 const G=sheets['Groundwork Review'].data;ok(n>=1&&G.some(r=>r[3]==='Gravel compacted under slab'&&r[1]==='4 Prep: Framing & Trades Rough In'&&r[5]==='g1'),'slab task listed with its current checklist');
 ok(!G.some(r=>r[3]==='Tape and mud'),'unrelated task not listed');ok(JSON.stringify(I)===before,'review list moves nothing');}
function SpreadsheetApp_(){return ctx.SpreadsheetApp.getActive();}
