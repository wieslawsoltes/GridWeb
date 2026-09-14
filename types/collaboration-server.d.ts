import {CollaborationStore,RoomStorage} from './collaboration.js';
export interface Request {readonly headers:Record<string,string|string[]|undefined>;readonly method?:string;readonly url?:string;}
export interface Principal {id:string;role:'editor'|'viewer';}
/** The returned instance is a node:http Server; this structural facade avoids a browser dependency on @types/node. */
export interface CollaborationHttpServer {
 readonly listening:boolean;
 listen(port:number,hostname:string,callback?:()=>void):this;
 address():{port:number;address:string;family:string}|string|null;
 on(name:string,callback:(...args:unknown[])=>void):this;
}
export interface ServerOptions {
 authorize:(request:Request,resource:{room:string;action:string})=>Principal|null|Promise<Principal|null>;
 storageDirectory?:string;storage?:RoomStorage;allowedOrigins?:string[];studioFile?:string|null;
 maxBodyBytes?:number;maxStreams?:number;maxRooms?:number;validationTimeoutMs?:number;
}
export function createCollaborationServer(options:ServerOptions):Promise<{server:CollaborationHttpServer;store:CollaborationStore;close():Promise<void>}>;
export function createFileStorage(directory:string):Promise<RoomStorage>;
export function bearerAuthorization(token:string,principal?:Partial<Principal>):(request:Request)=>Principal|null;
