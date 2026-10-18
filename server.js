'use strict';
// Zero-dependency static web server for local use or Railway.
const http=require('http');
const fs=require('fs');
const path=require('path');
const root=__dirname;
const types={'.html':'text/html; charset=utf-8','.js':'text/javascript; charset=utf-8','.css':'text/css; charset=utf-8','.jpg':'image/jpeg','.jpeg':'image/jpeg','.png':'image/png','.svg':'image/svg+xml','.xlsx':'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet','.txt':'text/plain; charset=utf-8','.docx':'application/vnd.openxmlformats-officedocument.wordprocessingml.document','.pdf':'application/pdf','.ico':'image/x-icon'};
const port=process.env.PORT||3000;
http.createServer((req,res)=>{
 const pathname=new URL(req.url,'http://localhost').pathname;
 const requested=decodeURIComponent(pathname);
 const filePath=path.resolve(root,'.'+requested,(requested==='/'?'index.html':''));
 if(filePath!==root&&!filePath.startsWith(root+path.sep)){res.writeHead(403);res.end('Forbidden');return;}
 fs.stat(filePath,(err,st)=>{
  if(err||!st.isFile()){res.writeHead(404,{'Content-Type':'text/plain'});res.end('Not found');return;}
  const extension=path.extname(filePath).toLowerCase();
  const headers={'Content-Type':types[extension]||'application/octet-stream','X-Content-Type-Options':'nosniff','Content-Length':st.size};
  if(extension==='.xlsx'){headers['Content-Disposition']='attachment; filename="Revenue Dashboard 1.xlsx"';}
  res.writeHead(200,headers);
  if(req.method==='HEAD'){res.end();return;}
  fs.createReadStream(filePath).pipe(res);
 });
}).listen(port,'0.0.0.0',()=>console.log(`CWT Power BI Training site running on port ${port}`));
