Office.onReady(function(){
  document.getElementById('insertBtn').addEventListener('click', buildInOutlook);
  document.getElementById('previewBtn').addEventListener('click', preview);
});

function value(id){ return document.getElementById(id).value.trim(); }
function esc(s){
  return String(s || '').replace(/[&<>"']/g, function(c){
    return {'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#039;'}[c];
  });
}
function textHtml(s){ return s ? esc(s).replace(/\n/g,'<br>') : ''; }
function p(s, cls){
  if(!s) return '';
  return '<p'+(cls?' class="'+cls+'"':'')+' style="font:18px/1.68 Garamond,Times New Roman,serif;margin:0 0 24px;">'+textHtml(s)+'</p>';
}
function image(url, alt){
  if(!url) return '';
  return '<img src="'+esc(url)+'" alt="'+esc(alt||'')+'" style="width:100%;height:auto;display:block;margin:0 0 9px;">';
}
function pleasureRows(){
  const data=[
    [value('pleasure1Label'),value('pleasure1')],
    [value('pleasure2Label'),value('pleasure2')],
    [value('pleasure3Label'),value('pleasure3')],
    [value('pleasure4Label'),value('pleasure4')],
    [value('pleasure5Label'),value('pleasure5')]
  ].filter(function(x){return x[0]||x[1]});
  return data.map(function(x){
    return '<tr><td style="vertical-align:top;padding:7px 18px 7px 0;width:105px;font:10px Arial,sans-serif;letter-spacing:1.2px;text-transform:uppercase;color:#777168;">'+esc(x[0])+'</td><td style="padding:7px 0;font:17px/1.45 Garamond,Times New Roman,serif;">'+textHtml(x[1])+'</td></tr>';
  }).join('');
}
function buildHtml(){
  const title=value('title') || 'The Pleasure Dispatch';
  const subtitle=value('subtitle');
  const meta=[value('edition'),value('date')].filter(Boolean).join(' · ');
  const motif='<div style="width:20px;height:20px;position:relative;margin:0 auto;"><span style="position:absolute;width:13px;height:13px;left:3px;top:0;border:1px solid #111;border-radius:50%;transform:rotate(45deg);"></span><span style="position:absolute;width:13px;height:13px;left:3px;top:7px;border:1px solid #111;border-radius:50%;transform:rotate(45deg);background:#fffdf8;"></span></div>';
  const cta=value('ctaUrl') ? '<a href="'+esc(value('ctaUrl'))+'" style="display:inline-block;padding:12px 18px;background:#111;color:#fff;text-decoration:none;font:10px Arial,sans-serif;letter-spacing:1.4px;">'+esc(value('ctaLabel')||'INQUIRE')+'</a>' : '';
  return '<div style="background:#f7f4ee;padding:28px 14px;color:#161616;">'
  +'<div style="max-width:700px;margin:0 auto;background:#fffdf8;padding:38px 34px 32px;">'
  +'<div style="font:11px Arial,sans-serif;letter-spacing:2px;font-weight:600;">FLRS GLOBAL</div>'
  +'<div style="font:15px/1.3 Garamond,Times New Roman,serif;font-style:italic;margin-top:7px;">From the Studio of Freddie L. Rankin II</div>'
  +'<div style="border-bottom:1px solid #111;margin-top:18px;padding-bottom:20px;">'
  +'<div style="font:10px Arial,sans-serif;letter-spacing:1.4px;font-weight:600;">THE PLEASURE DISPATCH</div>'
  +'<div style="font:11px Arial,sans-serif;letter-spacing:1px;color:#777168;margin-top:8px;">'+esc(meta)+'</div>'
  +'<div style="font:40px/1.03 Garamond,Times New Roman,serif;margin-top:16px;">'+esc(title)+'</div>'
  +(subtitle?'<div style="font:21px/1.2 Garamond,Times New Roman,serif;font-style:italic;margin-top:7px;">'+esc(subtitle)+'</div>':'')
  +'</div>'
  +'<div style="padding:20px 0;border-bottom:1px solid #cfc9bf;">'+motif+'</div>'
  +'<div style="padding-top:28px;">'
  +'<div style="font:10px Arial,sans-serif;letter-spacing:1.4px;font-weight:600;color:#777168;margin-bottom:12px;">01 — A REFLECTION</div>'
  +p(value('reflection'))
  +'<div style="font:10px Arial,sans-serif;letter-spacing:1.4px;font-weight:600;color:#777168;margin:34px 0 12px;">02 — THE WORK</div>'
  +image(value('workImage'),value('workTitle'))
  +(value('workTitle')?'<div style="font:25px/1.1 Garamond,Times New Roman,serif;margin:10px 0 5px;">'+esc(value('workTitle'))+'</div>':'')
  +(value('workMeta')?'<div style="font:10px Arial,sans-serif;letter-spacing:1px;color:#777168;text-transform:uppercase;margin-bottom:13px;">'+esc(value('workMeta'))+'</div>':'')
  +p(value('workText'))
  +'<div style="font:10px Arial,sans-serif;letter-spacing:1.4px;font-weight:600;color:#777168;margin:34px 0 12px;">03 — STUDIO NOTES</div>'
  +image(value('studioImage'),value('studioTitle'))
  +(value('studioTitle')?'<div style="font:25px/1.1 Garamond,Times New Roman,serif;margin:10px 0 7px;">'+esc(value('studioTitle'))+'</div>':'')
  +p(value('studioText'))
  +'<div style="font:10px Arial,sans-serif;letter-spacing:1.4px;font-weight:600;color:#777168;margin:34px 0 8px;">04 — PLEASURE NOTES</div>'
  +'<div style="font:17px/1.35 Garamond,Times New Roman,serif;font-style:italic;margin-bottom:12px;">An offering of what has held my attention.</div>'
  +'<table cellpadding="0" cellspacing="0" style="border-collapse:collapse;width:100%;margin-bottom:6px;">'+pleasureRows()+'</table>'
  +'<div style="font:10px Arial,sans-serif;letter-spacing:1.4px;font-weight:600;color:#777168;margin:34px 0 12px;">05 — AN INVITATION</div>'
  +(value('inviteTitle')?'<div style="font:25px/1.1 Garamond,Times New Roman,serif;margin-bottom:9px;">'+esc(value('inviteTitle'))+'</div>':'')
  +p(value('inviteText'))+cta
  +'<div style="font:10px Arial,sans-serif;letter-spacing:1.4px;font-weight:600;color:#777168;margin:34px 0 12px;">06 — CLOSING</div>'
  +p(value('closing'))
  +(value('question')?'<p style="font:24px/1.35 Garamond,Times New Roman,serif;margin:25px 0 35px;">'+esc(value('question'))+'</p>':'')
  +'</div>'
  +'<div style="border-top:1px solid #111;padding-top:18px;text-align:center;">'+motif+'<div style="font:10px Arial,sans-serif;letter-spacing:1px;margin-top:13px;">THE PLEASURE DISPATCH · BY FLRS GLOBAL</div></div>'
  +'</div></div>';
}
function subject(){
  return 'The Pleasure Dispatch — '+(value('edition')||'No. 001')+': '+(value('title')||'A Reflection');
}
function buildInOutlook(){
  const item=Office.context.mailbox.item;
  if(!item||!item.body){ setStatus('Open a new Outlook message to use the composer.'); return; }
  item.subject.setAsync(subject(),function(sr){
    if(sr.status!==Office.AsyncResultStatus.Succeeded){ setStatus('Could not set the subject.'); return; }
    item.body.setAsync(buildHtml(),{coercionType:Office.CoercionType.Html},function(br){
      setStatus(br.status===Office.AsyncResultStatus.Succeeded ? 'Dispatch built in Outlook. Review it before sending.' : 'Could not insert the Dispatch.');
    });
  });
}
function preview(){
  const w=window.open('','_blank');
  if(!w){setStatus('Preview was blocked by the browser.');return;}
  w.document.open();w.document.write(buildHtml());w.document.close();
}
function setStatus(s){document.getElementById('status').textContent=s;}
