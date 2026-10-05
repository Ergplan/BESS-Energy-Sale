// Step 3 · Run the model's own JS headlessly and write export.json (own-solar dispatch + reference results) for the Excel.
// Run after build.py:  node source/export_dispatch.js
const fs = require('fs'), path = require('path');
const s = fs.readFileSync(path.join(__dirname, '.model.js'), 'utf8');
const core = s.slice(0, s.indexOf('/* ---------- render: sidebar'));
const v2 = s.slice(s.indexOf('const T2 = {'), s.indexOf('function update2(){'));
const test = `
const base={solarCapex:2.05,bessCapex:0.98,landRate:35000,landMode:'lease',landEsc:5,cuf:30,auxMW:1,surplus:'sell',dod:90,rte:90,deg:1.5,bessAux:1,lineLoss:1,istsLoss:3,avail:99,augYear:12,augPct:20,stoa:586,lta:3.96,fee:15,life:25,tax:25.17,esc:0,omSolar:5,omBess:0.5,omEsc:3,solDeg:0.5,solarSize:100,debtPct:70,kd:9.5,tenor:15,ke:14,boostOn:false,boostDays:20,boostStart:1,boostPrice:15,eqOn:true,eqCost:20,eqYears:2};
const out={dispatch:{}, results:[], catl:{}};
S.win={2:{c:44,d:76},4:{c:38,d:72}};
for (const pk of ['P7','P8','P12']){ S.plant=pk;
  for (const h of [2,4]){
    for (const bo of [0,1]){ S.a=Object.assign({},base,{boostOn:!!bo}); const R=run(scOf(h,'solar','stoa'));
      out.dispatch[pk+'|'+h+'|'+bo]={bessMWh:R.years.map(y=>y.bessMWh), bessRev:R.years.map(y=>y.bessRev), expMWh:R.years.map(y=>y.expMWh), expRev:R.years.map(y=>y.expRev), fullDays:R.y1.fullDays}; }
    for (const [src,tx,lab] of [['solar','stoa','Own solar'],['grid','stoa','IEX + T-GNA'],['grid','lta','IEX + GNA']]) for (const bo of [0,1]) for (const eq of [0,1]){
      S.a=Object.assign({},base,{boostOn:!!bo, eqOn:!!eq}); const R=run(scOf(h,src,tx));
      out.results.push({pk,h,lab,bo,eq,irr:R.irr,irrPre:R.irrPre,eirr:R.eirr,npv:R.npv/1e7,capex:R.capex/1e7,dscr:R.dscrMin,lcos:R.lcos/1000,name:R.z.name,ac:R.z.ac,sold:R.y1.bessMWh,ebitda1:R.rows[0].ebitda/1e7}); }
  }
}
S.a=Object.assign({},base);
const C={capex:1.05,life:12000,eol:70,cal:0.3,dod:90,rte:90,aux:1,avail:99,rep:35};
for (const pk of ['P7','P8','P12']) for (const h of [2,4]) for (const tx of ['stoa','lta']){ S.plant=pk; T2.dur=h; T2.tx=tx; T2.buy='DAM'; T2.sell='DAM';
  const z1=size2(h,1,C), z2=size2(h,2,C); const WB=winSums('DAM',z2.nc), WS=winSums('DAM',z2.nb), WB1= z1.nc===z2.nc?WB:winSums('DAM',z1.nc);
  const w2=best2(WB,WS,z2), w1=best1(WB1,WS,z1); const R1=run2(1,w1,C,WB1,WS), R2=run2(2,w2,C,WB,WS);
  const pick=R=>({irr:R.irr,eirr:R.eirr,npv:R.npv/1e7,capex:R.capex/1e7,reps:R.reps,dscr:R.dscrMin});
  out.catl[pk+'|'+h+'|'+tx]={w1,w2,r1:pick(R1),r2:pick(R2)}; }
require('fs').writeFileSync(require('path').join(__dirname,'export.json'), JSON.stringify(out));
console.log('export.json written');`;
eval(core + v2 + test);
