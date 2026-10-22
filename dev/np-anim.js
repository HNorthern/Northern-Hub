(function(){
  if(window.__npAnim) return; window.__npAnim=true;
  var reduce=window.matchMedia&&matchMedia('(prefers-reduced-motion: reduce)').matches;
  var EASE='cubic-bezier(.2,.8,.2,1)', seen=new WeakSet(), lastPlay=new WeakMap();
  function num(str){ var m=String(str).match(/-?[\d,]*\.?\d+/); return m?m[0]:null; }
  function countUp(el,delay){
    var tn=null; for(var i=0;i<el.childNodes.length;i++){ if(el.childNodes[i].nodeType===3&&/\d/.test(el.childNodes[i].nodeValue)){tn=el.childNodes[i];break;} }
    if(!tn) return; if(tn.__npRun) return;
    var target=tn.nodeValue, n=num(target); if(!n) return;
    var dec=(n.split('.')[1]||'').length, commas=n.indexOf(',')>=0, end=parseFloat(n.replace(/,/g,'')); if(!isFinite(end)||end===0) return;
    var dur=1000, t0=performance.now()+(delay||0), last=null; tn.__npRun=true;
    function fmt(v){ var s=Math.abs(v).toFixed(dec); if(commas){ var p=s.split('.'); p[0]=p[0].replace(/\B(?=(\d{3})+(?!\d))/g,','); s=p.join('.'); } return (v<0?'-':'')+s; }
    function step(now){
      if(last!==null&&tn.nodeValue!==last){ tn.__npRun=false; return; }
      var k=Math.max(0,Math.min(1,(now-t0)/dur)), e=1-Math.pow(1-k,3);
      if(k>=1){ tn.nodeValue=target; tn.__npRun=false; return; }
      last=target.replace(n,fmt(end*e)); tn.nodeValue=last; requestAnimationFrame(step);
    }
    requestAnimationFrame(step);
  }
  function play(root,force){
    if(reduce) return;
    var now=performance.now(); if(!force&&now-(lastPlay.get(root)||0)<1400) return; lastPlay.set(root,now);
    var gx=root.querySelectorAll('[data-np-grow="x"]'); gx.forEach(function(el,i){ el.animate([{transform:'scaleX(0)',transformOrigin:'left center'},{transform:'scaleX(1)',transformOrigin:'left center'}],{duration:850,delay:Math.min(i,6)*90,easing:EASE,fill:'backwards'}); });
    var gy=root.querySelectorAll('[data-np-grow="y"]'); gy.forEach(function(el,i){ el.animate([{transform:'scaleY(0)',transformOrigin:'center bottom'},{transform:'scaleY(1)',transformOrigin:'center bottom'}],{duration:750,delay:Math.min(i,24)*(gy.length>8?25:110),easing:EASE,fill:'backwards'}); });
    root.querySelectorAll('[data-np-sweep]').forEach(function(el){ el.animate([{'--np-sw':'0deg'},{'--np-sw':'360deg'}],{duration:1100,easing:EASE}); el.parentNode.animate([{transform:'rotate(-90deg) scale(.94)'},{transform:'rotate(0) scale(1)'}],{duration:1100,easing:EASE}); });
    root.querySelectorAll('[data-np-draw]').forEach(function(el){ el.animate([{clipPath:'inset(0 100% 0 0)'},{clipPath:'inset(0 0 0 0)'}],{duration:1300,easing:EASE}); });
    root.querySelectorAll('[data-np-count]').forEach(function(el,i){ countUp(el,Math.min(i,6)*60); });
  }
  function scan(){
    document.querySelectorAll('[data-np-card]').forEach(function(c,i){ if(!seen.has(c)&&c.offsetParent){ seen.add(c); setTimeout(function(){play(c,true);},Math.min(i,8)*70); } });
  }
  document.addEventListener('mouseover',function(e){
    var c=e.target.closest&&e.target.closest('[data-np-card]'); if(!c) return;
    if(e.relatedTarget&&c.contains(e.relatedTarget)) return; play(c,false);
  },true);
  new MutationObserver(function(){ clearTimeout(window.__npScanT); window.__npScanT=setTimeout(scan,60); }).observe(document.documentElement,{childList:true,subtree:true});
  setTimeout(scan,300);
})();
