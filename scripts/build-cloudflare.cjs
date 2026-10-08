// Copy only reviewed public runtime assets; keep development documents private.
const fs=require('node:fs'),path=require('node:path');
const root=path.resolve(__dirname,'..');
const config=JSON.parse(fs.readFileSync(path.join(root,'cloudflare-public.json'),'utf8'));
const out=path.join(root,'dist');
if(path.dirname(out)!==root||path.basename(out)!=='dist')throw Error('Invalid output directory');
fs.rmSync(out,{recursive:true,force:true});fs.mkdirSync(out,{recursive:true});
for(const file of config.files){
 const from=path.resolve(root,file),to=path.resolve(out,file);
 if(!from.startsWith(root+path.sep)||!to.startsWith(out+path.sep))throw Error('Invalid public path: '+file);
 if(!fs.existsSync(from))throw Error('Missing public asset: '+file);
 fs.mkdirSync(path.dirname(to),{recursive:true});fs.cpSync(from,to,{recursive:true});
}
if(config.templates){
 for(const name of fs.readdirSync(path.join(root,'templates')))if(/\.(xlsx|png|json)$/i.test(name)){
  fs.mkdirSync(path.join(out,'templates'),{recursive:true});fs.copyFileSync(path.join(root,'templates',name),path.join(out,'templates',name));
 }
 for(const file of ['holdings-v0.1.source.html','management-v0.6.2.source.html','migration-v0.6.1.source.html']){
  const target=path.join(out,'site-src',file),text=fs.readFileSync(target,'utf8');let count=0;
  const adapted=text.replace(/(['"])https:\/\/raw\.githubusercontent\.com\/kirakein-gif\/ddalkkak-records\/main\/templates(\/?)\1/g,(_,quote,slash)=>{count++;return "(location.origin+'/templates"+slash+"')";});
  if(count!==1)throw Error('Template adapter needs review: '+file+' ('+count+' matches)');
  fs.writeFileSync(target,adapted.replaceAll('표준서식을 GitHub에서 확인하는 중','표준서식을 확인하는 중').replaceAll('GitHub 템플릿 다운로드 실패','표준서식 다운로드 실패'));
 }
}
if(config.repo==='ddalkkak-cart-request'){
 const target=path.join(out,'cart-extract.js'),text=fs.readFileSync(target,'utf8');
 const old="var appUrl = 'https://kirakein-gif.github.io/ddalkkak-cart-request/';";
 if(!text.includes(old))throw Error('Cart fallback adapter needs review');
 fs.writeFileSync(target,text.replace(old,'var appUrl = '+JSON.stringify(config.publicUrl)+';'));
}
fs.writeFileSync(path.join(out,'_headers'),'/version.json\n  Cache-Control: no-cache\n/deployment.json\n  Cache-Control: no-cache\n/cart-extract.js\n  Cache-Control: public, max-age=300\n/*\n  X-Content-Type-Options: nosniff\n  Referrer-Policy: strict-origin-when-cross-origin\n');
fs.writeFileSync(path.join(out,'404.html'),'<!doctype html><html lang="ko"><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>페이지를 찾을 수 없습니다</title><body style="font-family:Malgun Gothic,sans-serif;padding:40px"><h1>페이지를 찾을 수 없습니다.</h1><p><a href="/">프로그램 첫 화면으로 이동 →</a></p></body></html>');
fs.writeFileSync(path.join(out,'deployment.json'),JSON.stringify({appVersion:config.version,hostingAdapterVersion:1,commit:process.env.CF_PAGES_COMMIT_SHA||null},null,2));
console.log('Cloudflare static build ready: '+config.repo+' → dist');
