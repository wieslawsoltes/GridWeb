#!/usr/bin/env node
import path from 'node:path';
import { createCollaborationServer, bearerAuthorization } from '../src/collaboration/server.js';
const token=process.env.GRIDWEB_COLLAB_TOKEN;
if(!token||token.length<24){console.error('Set GRIDWEB_COLLAB_TOKEN to a random secret with at least 24 characters. Tokens are never stored in workbooks.');process.exit(1);}
const port=Number(process.env.PORT??8099);if(!Number.isInteger(port)||port<1||port>65535)throw new Error('Invalid PORT');
const host=process.env.HOST??'127.0.0.1';
const app=await createCollaborationServer({authorize:bearerAuthorization(token),storageDirectory:process.env.GRIDWEB_DATA??'./gridweb-data',allowedOrigins:(process.env.GRIDWEB_ORIGINS??`http://127.0.0.1:${port},http://localhost:${port}`).split(',').filter(Boolean),studioFile:path.resolve(import.meta.dirname,'../dist/GridWeb-standalone.html')});
app.server.listen(port,host,()=>console.log(`GridWeb collaboration listening on ${host}:${port}. Open /studio. Use TLS for non-loopback access.`));
for(const signal of ['SIGINT','SIGTERM'])process.once(signal,()=>app.close().then(()=>process.exit(0)));
