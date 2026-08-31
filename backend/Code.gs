const REPO_OWNER = 'flrsgloba';
const REPO_NAME = 'dispatch.flrsglobal.com';
const REPO_BRANCH = 'main';
const DATA_PATH = 'data/dispatches.json';

/*
 * The Drive bridge and publisher are intentionally the same Apps Script
 * deployment. PUBLISH_KEY is the shared secret used by the Outlook add-in.
 */
const FOLDER_ID = '1Cv9RLIjU0Qa3N_0-EYYXJYBhDHrQoeAg';
const DISPATCH_BASE_URL = 'https://dispatch.flrsglobal.com';

function props_(){ return PropertiesService.getScriptProperties(); }
function json_(data){ return ContentService.createTextOutput(JSON.stringify(data)).setMimeType(ContentService.MimeType.JSON); }
function getPublishKey_(){ const key=props_().getProperty('PUBLISH_KEY'); if(!key) throw new Error('PUBLISH_KEY is not configured.'); return key; }
function getFolder(){ return DriveApp.getFolderById(FOLDER_ID); }

function doGet(e){
  try{
    const action = e && e.parameter && e.parameter.action ? String(e.parameter.action).toLowerCase().trim() : 'status';
    if(action === 'list') return json_(readFeed_());
    if(action === 'get') return json_(getEdition_(e.parameter.edition));
    if(action === 'status') return json_({status:'ok',service:'The Pleasure Dispatch publishing and Drive bridge',folderName:getFolder().getName()});
    return json_({status:'error',code:'UNKNOWN_ACTION',message:'Unknown GET action: '+action});
  }catch(err){ return json_({status:'error',code:'SERVER_ERROR',message:err.message||String(err)}); }
}

function doPost(e){
  try{
    const body = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    const action = body.action ? String(body.action).toLowerCase().trim() : '';

    if(!body.publishKey || String(body.publishKey) !== getPublishKey_()){
      return json_({status:'error',code:'INVALID_PUBLISH_KEY',message:'Publishing key rejected.'});
    }

    if(action === 'upload') return uploadImage_(body);
    if(action === 'publish') return publishDispatch_(body);
    return json_({status:'error',code:'UNKNOWN_ACTION',message:'Unknown action: '+(action||'[empty]')});
  }catch(err){ return json_({status:'error',code:'SERVER_ERROR',message:err.message||String(err)}); }
}

/* =========================================================
   PUBLISH DISPATCH → GITHUB
========================================================= */
function publishDispatch_(body){
  validate_(body);

  const feed = readFeed_();
  const record = {
    edition: String(body.edition),
    editionLabel: String(body.editionLabel || body.edition),
    date: String(body.date || ''),
    title: String(body.title || ''),
    subtitle: String(body.subtitle || ''),
    subject: String(body.subject || ''),
    html: String(body.html),
    searchText: String(body.searchText || ''),
    publishedAt: new Date().toISOString()
  };

  const existing = feed.dispatches.findIndex(function(d){ return String(d.edition) === record.edition; });
  if(existing >= 0) feed.dispatches[existing] = record;
  else feed.dispatches.push(record);

  feed.dispatches.sort(function(a,b){ return editionValue_(b.edition)-editionValue_(a.edition); });
  feed.updatedAt = new Date().toISOString();
  writeFeed_(feed);

  return json_({
    status:'published',
    edition:record.edition,
    title:record.title,
    slug:cleanSlug_(record.edition+'-'+record.title),
    updatedAt:feed.updatedAt,
    dispatchUrl:DISPATCH_BASE_URL+'/dispatch/'+encodeURIComponent(cleanSlug_(record.edition+'-'+record.title))
  });
}

function validate_(b){
  if(!b.edition) throw new Error('Edition is required.');
  if(!b.title) throw new Error('Title is required.');
  if(!b.html) throw new Error('HTML is required.');
  if(String(b.html).length > 900000) throw new Error('Dispatch is too large.');
}

function editionValue_(e){ const n=parseFloat(String(e).replace(/[^0-9.]/g,'')); return isNaN(n)?-1:n; }
function cleanSlug_(value){ return String(value||'').toLowerCase().trim().replace(/['"]/g,'').replace(/[^a-z0-9]+/g,'-').replace(/^-+|-+$/g,'').substring(0,100); }

/* =========================================================
   GITHUB
========================================================= */
function readFeed_(){
  try{
    const text=githubGet_(DATA_PATH);
    const data=JSON.parse(text||'{}');
    data.dispatches=Array.isArray(data.dispatches)?data.dispatches:[];
    return data;
  }catch(err){
    if(String(err.message).indexOf('404')>=0) return {version:1,updatedAt:null,dispatches:[]};
    throw err;
  }
}

function getEdition_(edition){
  const feed=readFeed_();
  return feed.dispatches.find(function(d){ return String(d.edition)===String(edition); }) || {status:'not_found',edition:String(edition||'')};
}

function writeFeed_(feed){
  const current=githubFile_(DATA_PATH);
  if(!current || !current.sha) throw new Error('GitHub data/dispatches.json could not be read before publishing.');
  githubPut_(DATA_PATH,JSON.stringify(feed,null,2),current.sha,'Publish The Pleasure Dispatch '+feed.dispatches[0].edition);
}

function githubFile_(path){
  const token=props_().getProperty('GITHUB_TOKEN');
  if(!token) throw new Error('GITHUB_TOKEN is not configured.');
  const url='https://api.github.com/repos/'+REPO_OWNER+'/'+REPO_NAME+'/contents/'+path+'?ref='+REPO_BRANCH;
  const r=UrlFetchApp.fetch(url,{method:'get',muteHttpExceptions:true,headers:{Authorization:'Bearer '+token,Accept:'application/vnd.github+json'}});
  if(r.getResponseCode()===404) return null;
  if(r.getResponseCode()>=300) throw new Error('GitHub read failed: HTTP '+r.getResponseCode());
  return JSON.parse(r.getContentText());
}

function githubGet_(path){
  const f=githubFile_(path);
  if(!f||!f.content) throw new Error('GitHub file not found: '+path);
  return Utilities.newBlob(Utilities.base64Decode(f.content.replace(/\n/g,''))).getDataAsString('UTF-8');
}

function githubPut_(path,content,sha,message){
  const token=props_().getProperty('GITHUB_TOKEN');
  if(!token) throw new Error('GITHUB_TOKEN is not configured.');
  const url='https://api.github.com/repos/'+REPO_OWNER+'/'+REPO_NAME+'/contents/'+path;
  const payload={message:message,content:Utilities.base64Encode(Utilities.newBlob(content).getBytes()),branch:REPO_BRANCH,sha:sha};
  const r=UrlFetchApp.fetch(url,{method:'put',muteHttpExceptions:true,contentType:'application/json',payload:JSON.stringify(payload),headers:{Authorization:'Bearer '+token,Accept:'application/vnd.github+json'}});
  if(r.getResponseCode()>=300) throw new Error('GitHub publish failed: HTTP '+r.getResponseCode()+' '+r.getContentText());
}

/* =========================================================
   DRIVE IMAGE UPLOAD
========================================================= */
function uploadImage_(data){
  if(!data.fileName) throw new Error('Missing fileName.');
  if(!data.mimeType) throw new Error('Missing mimeType.');
  if(!data.fileContent) throw new Error('Missing fileContent.');

  const mimeType=String(data.mimeType).toLowerCase().trim();
  if(!/^image\/(jpeg|jpg|png|webp)$/i.test(mimeType)) throw new Error('Only JPG, PNG, and WebP images are supported.');

  let bytes;
  try{ bytes=Utilities.base64Decode(String(data.fileContent)); }
  catch(err){ throw new Error('The image data could not be decoded.'); }
  if(!bytes || !bytes.length) throw new Error('The decoded image is empty.');

  const file=getFolder().createFile(Utilities.newBlob(bytes,mimeType,cleanFileName_(data.fileName)));
  let publicAccess=false;
  try{ file.setSharing(DriveApp.Access.ANYONE_WITH_LINK,DriveApp.Permission.VIEW); publicAccess=true; }catch(err){ publicAccess=false; }

  const fileId=file.getId();
  const imageUrl='https://drive.google.com/thumbnail?id='+encodeURIComponent(fileId)+'&sz=w1800';
  const driveUrl='https://drive.google.com/file/d/'+encodeURIComponent(fileId)+'/view';

  return json_({status:'created',fileId:fileId,fileName:file.getName(),mimeType:file.getMimeType(),sizeBytes:file.getSize(),driveUrl:driveUrl,publicAccess:publicAccess,imageUrl:imageUrl});
}

function cleanFileName_(name){ return String(name||'image').replace(/[\\/:*?"<>|#%{}\[\]]/g,'_').replace(/\s+/g,'_').replace(/[^a-zA-Z0-9._-]/g,'_').substring(0,150); }

/* =========================================================
   TEST
========================================================= */
function testPublishBridge(){
  const feed=readFeed_();
  Logger.log('GitHub feed contains '+feed.dispatches.length+' dispatch(es).');
  return json_({status:'success',dispatchCount:feed.dispatches.length,updatedAt:feed.updatedAt});
}
