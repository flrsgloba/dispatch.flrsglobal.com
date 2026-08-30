const REPO_OWNER = 'flrsgloba';
const REPO_NAME = 'dispatch.flrsglobal.com';
const REPO_BRANCH = 'main';
const DATA_PATH = 'data/dispatches.json';

function props_(){ return PropertiesService.getScriptProperties(); }
function json_(data){ return ContentService.createTextOutput(JSON.stringify(data)).setMimeType(ContentService.MimeType.JSON); }

function doGet(e){
  const action = e && e.parameter && e.parameter.action;
  if(action === 'list') return json_(readFeed_());
  if(action === 'get') return json_(getEdition_(e.parameter.edition));
  return json_({status:'ok',service:'The Pleasure Dispatch publishing bridge'});
}

function doPost(e){
  try{
    const body = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    if(body.action !== 'publish') return json_({status:'error',message:'Unknown action.'});
    if(body.secret !== props_().getProperty('PUBLISH_SECRET')) return json_({status:'error',message:'Unauthorized.'});
    validate_(body);
    const feed = readFeed_();
    const record = {edition:String(body.edition),editionLabel:String(body.editionLabel||body.edition),date:String(body.date||''),title:String(body.title||''),subtitle:String(body.subtitle||''),subject:String(body.subject||''),html:String(body.html),searchText:String(body.searchText||''),publishedAt:new Date().toISOString()};
    const existing = feed.dispatches.findIndex(d=>String(d.edition)===record.edition);
    if(existing>=0) feed.dispatches[existing]=record; else feed.dispatches.push(record);
    feed.dispatches.sort((a,b)=>editionValue_(b.edition)-editionValue_(a.edition));
    feed.updatedAt=new Date().toISOString();
    writeFeed_(feed);
    return json_({status:'published',edition:record.edition,updatedAt:feed.updatedAt});
  }catch(err){ return json_({status:'error',message:err.message||String(err)}); }
}

function validate_(b){ if(!b.edition)throw new Error('Edition is required.'); if(!b.title)throw new Error('Title is required.'); if(!b.html)throw new Error('HTML is required.'); if(b.html.length>900000)throw new Error('Dispatch is too large.'); }
function editionValue_(e){const n=parseFloat(String(e).replace(/[^0-9.]/g,''));return isNaN(n)?-1:n;}
function readFeed_(){try{const text=githubGet_(DATA_PATH);const data=JSON.parse(text||'{}');data.dispatches=Array.isArray(data.dispatches)?data.dispatches:[];return data;}catch(err){if(String(err.message).indexOf('404')>=0)return{version:1,updatedAt:null,dispatches:[]};throw err;}}
function getEdition_(edition){const feed=readFeed_();return feed.dispatches.find(d=>String(d.edition)===String(edition))||{status:'not_found',edition:String(edition||'')};}
function writeFeed_(feed){githubPut_(DATA_PATH,JSON.stringify(feed,null,2),githubFile_(DATA_PATH).sha,'Publish The Pleasure Dispatch '+feed.dispatches[0].edition);}
function githubFile_(path){const token=props_().getProperty('GITHUB_TOKEN');if(!token)throw new Error('GITHUB_TOKEN is not configured.');const url='https://api.github.com/repos/'+REPO_OWNER+'/'+REPO_NAME+'/contents/'+path+'?ref='+REPO_BRANCH;const r=UrlFetchApp.fetch(url,{method:'get',muteHttpExceptions:true,headers:{Authorization:'Bearer '+token,Accept:'application/vnd.github+json'}});if(r.getResponseCode()===404)return null;if(r.getResponseCode()>=300)throw new Error('GitHub read failed: HTTP '+r.getResponseCode());return JSON.parse(r.getContentText());}
function githubGet_(path){const f=githubFile_(path);if(!f||!f.content)throw new Error('GitHub file not found: '+path);return Utilities.newBlob(Utilities.base64Decode(f.content.replace(/\n/g,''))).getDataAsString('UTF-8');}
function githubPut_(path,content,sha,message){const token=props_().getProperty('GITHUB_TOKEN');const url='https://api.github.com/repos/'+REPO_OWNER+'/'+REPO_NAME+'/contents/'+path;const payload={message:message,content:Utilities.base64Encode(Utilities.newBlob(content).getBytes()),branch:REPO_BRANCH,sha:sha};const r=UrlFetchApp.fetch(url,{method:'put',muteHttpExceptions:true,contentType:'application/json',payload:JSON.stringify(payload),headers:{Authorization:'Bearer '+token,Accept:'application/vnd.github+json'}});if(r.getResponseCode()>=300)throw new Error('GitHub publish failed: HTTP '+r.getResponseCode()+' '+r.getContentText());}
