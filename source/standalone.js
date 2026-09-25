<script>
/* Standalone: two cycles a day · CATL 12,000-cycle · IEX only */
(function(){
  document.querySelector('.viewtabs').hidden = true;
  document.querySelector('.brand span').textContent = 'CATL 12,000-cycle · IEX only';
  document.querySelector('.brand b').textContent = 'Two-Cycle Arbitrage';
  document.querySelector('.keyin p').textContent = 'The BESS is sized to export this MW for the whole discharge window, in each cycle.';
  // move the grid, charge and finance inputs this scenario uses into its own card
  const groups = [
    ['Losses', ['a_lineLoss','a_istsLoss']],
    ['Transmission & exchange', ['a_stoa','a_lta','a_fee']],
    ['Finance & O&M', ['a_life','a_tax','a_esc','a_omBess','a_omEsc']],
    ['Financing & WACC', ['a_debtPct','a_kd','a_tenor','a_ke']],
  ];
  const sec = document.createElement('section'); sec.className = 'card';
  sec.innerHTML = '<div class="card-h"><div><span class="lbl">Inputs</span><h2>Grid losses, charges and finance</h2></div></div><div class="assm"></div>';
  const box = sec.querySelector('.assm');
  for (const [name, ids] of groups){
    const g = document.createElement('div'); g.className = 'grp'; const h = document.createElement('h3'); h.textContent = name; g.appendChild(h);
    ids.forEach(id => { const f = document.getElementById(id).closest('.field'); g.appendChild(f); });
    if (name.startsWith('Transmission')){ const p = document.createElement('p'); p.className = 'hint'; p.style.margin = '0';
      p.textContent = 'No CTU charges on selling. T-GNA ₹586/MWh and GNA ₹3.96 lakh/MW/month are derived from the SRPC RTA for May 2026 (Karnataka drawee DICs) using Reg. 11 of the CERC Sharing Regulations 2020. The MoP ISTS waiver covers only co-located BESS, so it does not apply here.'; g.appendChild(p); }
    box.appendChild(g);
  }
  const v2 = document.getElementById('view2'); v2.insertBefore(sec, v2.lastElementChild.nextSibling);
  document.querySelectorAll('#view2 .hint').forEach(p => { if (p.textContent.includes('Solar + BESS tab')) p.textContent = p.textContent.replace('come from the Assumptions on the Solar + BESS tab', 'are set in the Grid losses, charges and finance card below').replace(' Solar is not used on this tab.', ' No solar is used.'); });
  window.renderTabs = function(){
    const t = document.getElementById('plantTabs'); t.innerHTML = ''; const C = readT2();
    for (const k in PLANTS){ const p = PLANTS[k], cur = S.plant; S.plant = k; const z = size2(T2.dur, 2, C); S.plant = cur;
      const b = document.createElement('button'); b.className = 'tab'; b.setAttribute('role','tab'); b.setAttribute('aria-selected', S.plant===k ? 'true':'false');
      b.innerHTML = `<span class="t1"><span class="code">${p.code}</span><span class="loc">${p.loc}</span></span><span class="conn">${inr(p.conn)}<small>MW night connectivity</small></span><span class="t3">BESS ${inr(z.name)} MWh · ${T2.dur}-hour · IEX only</span>`;
      b.onclick = () => { S.plant = k; update2(); }; t.appendChild(b); }
    document.getElementById('connInput').value = PLANTS[S.plant].conn;
  };
  setView(2);
})();
</script>
