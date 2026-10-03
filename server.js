// Optional local development bridge. Production runs the same handler in Azure.
const express = require('express');
const cors = require('cors');
const proxy = require('./api/sharepoint');
const app = express();
app.disable('x-powered-by');
app.use(cors({origin:'https://localhost:3000',methods:['GET','POST'],allowedHeaders:['Content-Type','X-NERD-SharePoint-Authorization']}));
app.use(express.json({limit:'1mb'}));
app.get('/api/health', (_req,res)=>res.json({status:'ok',service:'NERD API'}));
app.post('/api/sharepoint', async (req,res,next)=>{
 try {
  const context={}; await proxy(context,req);
  res.status(context.res.status).set(context.res.headers).json(context.res.body);
 } catch(error) { next(error); }
});
app.use((error,_req,res,_next)=>res.status(error.status===413?413:500).json({error:'The local bridge could not process this request.'}));
app.listen(3001,'127.0.0.1',()=>console.log('NERD development bridge is listening on loopback port 3001.'));
