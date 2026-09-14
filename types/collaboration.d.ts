import type {Workbook, EventSource, WorkbookDocument} from './index.js';
export interface RoomSnapshot { revision:number; epoch:number; document:WorkbookDocument; }
export interface CommitResult extends RoomSnapshot { appliedRevision:number; duplicate:boolean; }
export type ConflictStrategy='local'|'remote';
export interface CellChange { sheet:string;cell:number;field:'content'|'style'|'comment';expected:unknown;value:unknown; }
export type DocumentChange={kind:'cells';changes:CellChange[]}|{kind:'replace';document:WorkbookDocument};
export type CommitRequest=DocumentChange&{id:string;clientId:string;baseRevision:number;epoch:number};
export class CollaborationError extends Error { readonly code:string;readonly details:unknown;constructor(code:string,message:string,details?:unknown); }
export interface CollaborationOptions {url:string;room:string;token:string;fetch?:typeof fetch;storage?:Pick<Storage,'getItem'|'setItem'>|null;storageKey?:string|null;clientId?:string|null;autoSync?:boolean;}
export interface Participant {actor:string;clientId:string;sheet:string;selection:string;}
export class CollaborationSession {
 constructor(workbook:Workbook,options:CollaborationOptions);
 readonly Workbook:Workbook;readonly Url:string;readonly Room:string;readonly ClientId:string;AutoSync:boolean;
 readonly Revision:number;readonly HasPendingChanges:boolean;readonly Status:string;
 readonly Conflict:null|{details:unknown;base:RoomSnapshot;local:WorkbookDocument;remote:RoomSnapshot};
 readonly Changed:EventSource<{status:string;revision:number;pending:boolean}>;
 readonly Conflicted:EventSource<NonNullable<CollaborationSession['Conflict']>>;
 readonly PresenceChanged:EventSource<Participant[]>;readonly Error:EventSource<Error>;
 Connect(options?:{mode?:'join'|'create'}):Promise<this>;Sync():Promise<boolean>;Resolve(strategy:ConflictStrategy):Promise<boolean>;SetPresence(sheetId:string,selection:string):Promise<unknown>;Dispose():void;
}
export interface RoomStorage {load(room:string):Promise<unknown>;save(room:string,state:unknown):Promise<void>;}
export class CollaborationStore {
 constructor(options?:Partial<RoomStorage>&{maxRooms?:number;notify?:(room:string,event:unknown)=>void});
 get(room:string):Promise<RoomSnapshot>;create(room:string,document:WorkbookDocument):Promise<RoomSnapshot>;commit(room:string,request:CommitRequest,actor?:string):Promise<CommitResult>;
}
export function diffDocuments(base:WorkbookDocument,local:WorkbookDocument,options?:{exclusive?:boolean}):DocumentChange;
export function applyChanges(document:WorkbookDocument,changes:CellChange[],options?:{resolve?:'reject'|ConflictStrategy}):{document:WorkbookDocument;conflicts:CellChange[]};
export function rebaseDocuments(base:WorkbookDocument,local:WorkbookDocument,remote:WorkbookDocument,options?:{exclusive?:boolean;resolve?:ConflictStrategy}):{document:WorkbookDocument;conflicts:CellChange[]};
