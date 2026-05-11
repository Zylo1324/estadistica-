const Excel = require('exceljs');

function normSInv(p){
  // Peter John Acklam approximation
  if (p<=0||p>=1) throw new Error('p');
  const a=[-3.969683028665376e+01,2.209460984245205e+02,-2.759285104469687e+02,1.383577518672690e+02,-3.066479806614716e+01,2.506628277459239e+00];
  const b=[-5.447609879822406e+01,1.615858368580409e+02,-1.556989798598866e+02,6.680131188771972e+01,-1.328068155288572e+01];
  const c=[-7.784894002430293e-03,-3.223964580411365e-01,-2.400758277161838e+00,-2.549732539343734e+00,4.374664141464968e+00,2.938163982698783e+00];
  const d=[7.784695709041462e-03,3.224671290700398e-01,2.445134137142996e+00,3.754408661907416e+00];
  const plow=0.02425, phigh=1-plow; let q,r;
  if(p<plow){q=Math.sqrt(-2*Math.log(p));return (((((c[0]*q+c[1])*q+c[2])*q+c[3])*q+c[4])*q+c[5])/((((d[0]*q+d[1])*q+d[2])*q+d[3])*q+1);}
  if(p>phigh){q=Math.sqrt(-2*Math.log(1-p));return -(((((c[0]*q+c[1])*q+c[2])*q+c[3])*q+c[4])*q+c[5])/((((d[0]*q+d[1])*q+d[2])*q+d[3])*q+1);}
  q=p-0.5; r=q*q; return (((((a[0]*r+a[1])*r+a[2])*r+a[3])*r+a[4])*r+a[5])*q/(((((b[0]*r+b[1])*r+b[2])*r+b[3])*r+b[4])*r+1);
}
function tApprox(p,df){const z=normSInv(p);return z+(z*z*z+z)/(4*df);} // rough

function getSample(ws){
  const vals=[];
  for(let r=8;r<500;r++){
    const b=ws.getCell(r,2).value; const a=ws.getCell(r,1).value;
    const v= typeof b==='number'?b:(typeof a==='number'?a:null);
    if(typeof v==='number') vals.push(v);
  }
  return vals;
}
const configs={
  1:{kind:'zmean',mu0:45,sigma:8,n:64,xbar:47,alpha:0.05,conf:0.95,tail:'right',theory:'Prueba Z para media con desviación poblacional conocida.'},
  2:{kind:'zmean',mu0:1620,sigma:120,n:49,xbar:1580,alpha:0.04,conf:0.96,tail:'left',theory:'Prueba Z para media (varianza poblacional conocida).'},
  3:{kind:'zmean_s',mu0:220,s:60,n:100,xbar:234,alpha:0.01,conf:0.98,tail:'right',theory:'Muestra grande (n≥30): aproximación Z usando s.'},
  4:{kind:'t_from_data',mu0:20,alpha:0.05,conf:0.90,tail:'two',theory:'Prueba t para media con muestra pequeña y normalidad.'},
  5:{kind:'t_from_data',mu0:18,alpha:0.06,conf:0.97,tail:'left',theory:'Prueba t para media con σ desconocida.'},
  6:{kind:'zmean',mu0:2100,sigma:300,n:64,xbar:null,alpha:0.05,conf:0.98,tail:'two',theory:'Prueba Z bilateral para media poblacional.'},
  7:{kind:'p_from_data',p0:0.70,n:120,alpha:0.04,conf:0.90,tail:'right',theory:'Prueba Z para una proporción poblacional.'},
  8:{kind:'p_from_data',p0:0.06,n:180,alpha:0.05,conf:0.92,tail:'left',theory:'Prueba Z para proporción con contraste unilateral.'},
  9:{kind:'p_count',p0:0.55,n:240,x:146,alpha:0.05,conf:0.95,tail:'two',theory:'Prueba bilateral para proporción.'},
  10:{kind:'theory_only',theory:'Definición de hipótesis y tipo de cola (media).'},
  11:{kind:'theory_only',theory:'Definición de hipótesis y tipo de prueba (proporción).'},
  12:{kind:'t_from_data',mu0:500,alpha:0.05,conf:0.95,tail:'two',theory:'Prueba t para media, n pequeño.'},
  13:{kind:'zmean',mu0:1500,sigma:280,n:36,xbar:1420,alpha:0.02,conf:0.94,tail:'left',theory:'Prueba Z unilateral izquierda.'},
  14:{kind:'zmean_s',mu0:11500,s:720,n:36,xbar:11680,alpha:0.03,conf:0.92,tail:'right',theory:'Muestra grande: aproximación Z para media.'},
  15:{kind:'t_from_data',mu0:100,alpha:0.01,conf:0.99,tail:'two',theory:'Calibración con prueba t bilateral.'},
  16:{kind:'p_count',p0:0.45,n:300,x:151,alpha:0.08,conf:0.95,tail:'right',theory:'Prueba de proporción unilateral derecha.'},
  17:{kind:'p_count',p0:0.12,n:220,x:31,alpha:0.05,conf:0.90,tail:'left',theory:'Prueba de proporción unilateral izquierda.'},
  18:{kind:'t_from_data',mu0:70,alpha:0.05,conf:0.95,tail:'right',theory:'Prueba t unilateral con n pequeño.'},
  19:{kind:'zmean',mu0:2.5,sigma:0.2,n:81,xbar:2.46,alpha:0.025,conf:0.95,tail:'left',theory:'Prueba Z para media con varianza conocida.'},
  20:{kind:'p_from_data',p0:0.30,n:150,alpha:0.05,conf:0.95,tail:'two',theory:'Prueba bilateral para proporción.'},
};

(async()=>{
const wb=new Excel.Workbook(); await wb.xlsx.readFile('UNIDAD 2/Practica_Unidad_2_20_ejercicios.xlsx');
for(let i=1;i<=20;i++){
  const ws=wb.getWorksheet(`Ejercicio ${i}`); const c=configs[i];
  ws.getCell('D7').value='Teoría aplicada:'; ws.getCell('E7').value=c.theory;
  if(c.kind==='theory_only'){
    ws.getCell('D9').value='H0:'; ws.getCell('E9').value='Según enunciado (igualdad)';
    ws.getCell('D10').value='H1:'; ws.getCell('E10').value='Según afirmación planteada';
    ws.getCell('D11').value='Tipo de prueba:'; ws.getCell('E11').value='Unilateral/Bilateral según H1';
    continue;
  }
  let xbar,s,n,pHat,mu0,p0,sigma,alpha,conf,tail,stat,zcrit,pval,dec,li,ls;
  alpha=c.alpha; conf=c.conf; tail=c.tail;
  if(c.kind.includes('mean')||c.kind.includes('t_')){mu0=c.mu0;}
  if(c.kind.includes('p_')){p0=c.p0;}
  if(c.kind==='t_from_data' || c.kind==='p_from_data'){
    const sample=getSample(ws); n=sample.length;
    if(c.kind==='t_from_data'){xbar=sample.reduce((a,b)=>a+b,0)/n; s=Math.sqrt(sample.reduce((a,b)=>a+(b-xbar)**2,0)/(n-1));}
    else {const ones=sample.filter(v=>v===1).length; pHat=ones/n;}
  }
  if(c.kind==='zmean' || c.kind==='zmean_s'){xbar=c.xbar ?? getSample(ws).reduce((a,b)=>a+b,0)/c.n; n=c.n; sigma=c.sigma; s=c.s;}
  if(c.kind==='p_count'){n=c.n; pHat=c.x/c.n;}

  const zalpha=tail==='two'?normSInv(1-alpha/2):normSInv(1-alpha);
  if(c.kind.startsWith('p_')){
    stat=(pHat-p0)/Math.sqrt((p0*(1-p0))/n);
    if(tail==='right') pval=1-0.5*(1+erf(stat/Math.SQRT2));
    else if(tail==='left') pval=0.5*(1+erf(stat/Math.SQRT2));
    else pval=2*(1-0.5*(1+erf(Math.abs(stat)/Math.SQRT2)));
    zcrit=tail==='two'?`${(-zalpha).toFixed(4)}, ${zalpha.toFixed(4)}`:zalpha.toFixed(4);
    li=pHat-normSInv(1-(1-conf)/2)*Math.sqrt((pHat*(1-pHat))/n);
    ls=pHat+normSInv(1-(1-conf)/2)*Math.sqrt((pHat*(1-pHat))/n);
  } else {
    const se=(sigma??s)/Math.sqrt(n);
    const useT = c.kind==='t_from_data';
    stat=(xbar-mu0)/se;
    const crit=useT?tApprox(1-(tail==='two'?alpha/2:alpha),n-1):zalpha;
    zcrit=tail==='two'?`${(-crit).toFixed(4)}, ${crit.toFixed(4)}`:crit.toFixed(4);
    const pz=0.5*(1+erf(stat/Math.SQRT2));
    if(tail==='right') pval=1-pz; else if(tail==='left') pval=pz; else pval=2*(1-0.5*(1+erf(Math.abs(stat)/Math.SQRT2)));
    const ccrit = useT?tApprox(1-(1-conf)/2,n-1):normSInv(1-(1-conf)/2);
    li=xbar-ccrit*se; ls=xbar+ccrit*se;
  }
  dec=pval<alpha?'Se rechaza H0':'No se rechaza H0';

  const rows=[
    ['a) Prueba de hipótesis',''],['H0', c.kind.startsWith('p_')?`p = ${p0}`:`μ = ${mu0}`],['H1', tail==='right'?(c.kind.startsWith('p_')?`p > ${p0}`:`μ > ${mu0}`):tail==='left'?(c.kind.startsWith('p_')?`p < ${p0}`:`μ < ${mu0}`):(c.kind.startsWith('p_')?`p ≠ ${p0}`:`μ ≠ ${mu0}`)],
    ['α',alpha],['Estadístico',stat],['Valor crítico',zcrit],['p-valor',pval],['Decisión',dec],
    ['b) Intervalo de confianza',`${(conf*100).toFixed(0)}%`],['Límite inferior',li],['Límite superior',ls],['Conclusión',`Con ${(conf*100).toFixed(0)}% de confianza, el parámetro está entre ${li.toFixed(4)} y ${ls.toFixed(4)}.`]
  ];
  let r=9; for(const [k,v] of rows){ws.getCell(r,4).value=k; ws.getCell(r,5).value=v; r++;}
}
await wb.xlsx.writeFile('UNIDAD 2/Practica_Unidad_2_20_ejercicios_RESUELTO.xlsx');
console.log('ok');
})();
function erf(x){const sign=x>=0?1:-1; x=Math.abs(x); const a1=0.254829592,a2=-0.284496736,a3=1.421413741,a4=-1.453152027,a5=1.061405429,p=0.3275911; const t=1/(1+p*x); const y=1-(((((a5*t+a4)*t)+a3)*t+a2)*t+a1)*t*Math.exp(-x*x); return sign*y;}
