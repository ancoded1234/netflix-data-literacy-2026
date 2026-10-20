'use strict';
(() => {
  const data = window.DECK_DATA;
  if (!data) { document.getElementById('slides').textContent = '분석 데이터 파일을 불러오지 못했습니다. assets/deck-data.js를 확인하세요.'; return; }
  const stats = data.statistics, all = stats.groups.All, movies = stats.groups.Movies, shows = stats.groups.Shows;
  const escape = v => String(v).replace(/[&<>"']/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
  const number = v => Number(v).toLocaleString('en-US');
  const percent = v => (v * 100).toFixed(2);
  const share = (p, group = all, denominator = 'A') => group.top_shares[String(p)]['share_' + denominator];
  const num = (value, path, suffix = '', cls = '') => `<span class="num ${cls}" data-stat="${escape(path)}">${escape(value)}</span>${suffix}`;
  const foot = (left, right='Netflix What We Watched · 2026 H1') => `<footer class="footnote"><span>${left}</span><span>${right}</span></footer>`;
  const header = (kicker, title, subtitle='') => `<div class="slide-head"><div class="eyebrow">${kicker}</div><h2>${title}</h2>${subtitle ? `<p class="subhead">${subtitle}</p>` : ''}</div>`;
  const slide = (title, contents, classes='') => `<section class="slide ${classes}" aria-label="${escape(title)}">${contents}</section>`;

  function concentrationSVG() {
    const x = p => 90 + p / 20 * 850, y = s => 390 - s * 330;
    const pts = [[0,0], ...[1,5,10,20].map(p => [p,share(p)])];
    const path = pts.map(([p,s],i) => `${i?'L':'M'}${x(p)},${y(s)}`).join(' ');
    const grid = [0,.25,.5,.75,1].map(v=>`<line x1="90" x2="940" y1="${y(v)}" y2="${y(v)}" stroke="#34343b"/><text x="68" y="${y(v)+6}" text-anchor="end" fill="#aaaab4" font-size="19">${v*100}%</text>`).join('');
    const marks = [1,5,10,20].map(p => {
      const s = share(p), row = all.top_shares[String(p)];
      return `<g data-p="${p}" data-share="${s}"><circle cx="${x(p)}" cy="${y(s)}" r="6" fill="#f34450"/><text x="${x(p)}" y="${y(s)-23}" text-anchor="middle" fill="#f5f5f7" font-weight="700" font-size="29">${percent(s)}%</text><text x="${x(p)}" y="425" text-anchor="middle" fill="#f5f5f7" font-size="22">상위 ${p}%</text><text x="${x(p)}" y="455" text-anchor="middle" fill="#aaaab4" font-size="18">${number(row.selected_rows)}행</text></g>`;
    }).join('');
    return `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 1020 520" role="img" aria-labelledby="concentration-title concentration-desc"><title id="concentration-title">전체 개별 항목의 누적 상위 시청 시간 점유율</title><desc id="concentration-desc">상위 1% 25.23%, 5% 52.82%, 10% 67.93%, 20% 82.64%. 분모는 Other 제외 개별 항목 시청 시간 합계입니다.</desc>${grid}<path d="${path} L940,390 L90,390 Z" fill="#e50914" opacity=".10"/><path d="${path}" fill="none" stroke="#f34450" stroke-width="4"/>${marks}<text x="515" y="504" fill="#aaaab4" text-anchor="middle" font-size="19">시청 시간 내림차순 상위 행 비율 (%)</text><text x="90" y="25" fill="#aaaab4" font-size="19">누적 시청 시간 점유율 (%)</text></svg>`;
  }

  const markup = window.createSlides({stats, escape, number, percent, share, num, foot, header, slide, concentrationSVG});
  document.getElementById('slides').innerHTML = markup.join('');
  const rankingTabs=[...document.querySelectorAll('[data-ranking]')];
  function selectRanking(group) {
    rankingTabs.forEach(tab=>{const selected=tab.dataset.ranking===group;tab.setAttribute('aria-selected',String(selected));tab.tabIndex=selected?0:-1;});
    document.querySelectorAll('.ranking-panel').forEach(panel=>{const selected=panel.id===`ranking-${group}`;panel.hidden=!selected;panel.inert=!selected;});
  }
  rankingTabs.forEach(tab=>{
    tab.addEventListener('click',()=>selectRanking(tab.dataset.ranking));
    tab.addEventListener('keydown',e=>{
      if(['ArrowLeft','ArrowRight','ArrowUp','ArrowDown'].includes(e.key)){
        e.preventDefault();e.stopPropagation();const other=rankingTabs.find(t=>t!==tab);selectRanking(other.dataset.ranking);other.focus();
      }
    });
  });
  selectRanking('Movies');
  const slides = [...document.querySelectorAll('.slide')];
  const prev = document.getElementById('prev'), next = document.getElementById('next');
  let current = 0;
  function fromHash() { const match = location.hash.match(/^#slide-(\d+)$/); return match ? Number(match[1])-1 : 0; }
  function show(index, updateHash=true) {
    current = Math.max(0, Math.min(slides.length-1, index));
    slides.forEach((el,i) => { const active = i===current; el.classList.toggle('is-active', active); el.setAttribute('aria-hidden', String(!active)); el.inert=!active; });
    document.getElementById('counter').textContent = `${String(current+1).padStart(2,'0')} / ${slides.length}`;
    prev.disabled = current===0; next.disabled = current===slides.length-1;
    document.getElementById('progress-fill').style.width = `${(current+1)/slides.length*100}%`;
    document.querySelector('.progress').setAttribute('aria-valuenow',String(current+1));
    document.title = `${current+1}/11 · ${slides[current].getAttribute('aria-label')} | 넷플릭스에서는 모두가 같은 것만 볼까?`;
    if(updateHash) history.replaceState(null,'',`#slide-${String(current+1).padStart(2,'0')}`);
  }
  function resize() {
    const scale = Math.min(window.innerWidth/1600,(window.innerHeight-66)/900);
    const wrap = document.getElementById('stage-wrap');
    wrap.style.width = `${1600*scale}px`; wrap.style.height = `${900*scale}px`;
    document.getElementById('stage').style.transform = `scale(${scale})`;
  }
  async function fullscreen() {
    try {
      if(document.fullscreenElement) await document.exitFullscreen();
      else await document.documentElement.requestFullscreen();
    } catch (error) { document.getElementById('fullscreen').title='브라우저 전체화면을 사용할 수 없습니다. F11 키를 사용하세요.'; }
  }
  prev.addEventListener('click',()=>show(current-1)); next.addEventListener('click',()=>show(current+1));
  document.getElementById('fullscreen').addEventListener('click',fullscreen);
  document.addEventListener('fullscreenchange',()=>{ const btn=document.getElementById('fullscreen'); btn.setAttribute('aria-pressed',String(!!document.fullscreenElement)); btn.textContent=document.fullscreenElement?'전체화면 종료 ⛶':'전체화면 ⛶'; resize(); });
  window.addEventListener('resize',resize); window.addEventListener('hashchange',()=>show(fromHash(),false));
  document.addEventListener('keydown', e=>{
    if(e.altKey||e.ctrlKey||e.metaKey||/INPUT|TEXTAREA|SELECT/.test(e.target.tagName)) return;
    if(e.key==='ArrowRight'||e.key==='PageDown'){e.preventDefault();show(current+1);}
    else if(e.key==='ArrowLeft'||e.key==='PageUp'){e.preventDefault();show(current-1);}
    else if(e.key==='Home'){e.preventDefault();show(0);}
    else if(e.key==='End'){e.preventDefault();show(slides.length-1);}
    else if(e.key.toLowerCase()==='f'){e.preventDefault();fullscreen();}
  });
  let touchStart=null;
  const wrap=document.getElementById('stage-wrap');
  wrap.addEventListener('touchstart',e=>{touchStart={x:e.changedTouches[0].clientX,y:e.changedTouches[0].clientY};},{passive:true});
  wrap.addEventListener('touchend',e=>{if(touchStart===null)return;const dx=e.changedTouches[0].clientX-touchStart.x,dy=e.changedTouches[0].clientY-touchStart.y;if(Math.abs(dx)>60&&Math.abs(dx)>Math.abs(dy))show(current+(dx<0?1:-1));touchStart=null;},{passive:true});
  wrap.addEventListener('touchcancel',()=>{touchStart=null;},{passive:true});
  show(fromHash()); resize();
  // Read-only state for reproducible browser QA.
  window.DECK_STATE = {get current(){return current;},get count(){return slides.length;}};
})();
