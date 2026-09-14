import {Workbook} from '@wieslawsoltes/gridweb';
import {CollaborationSession,CollaborationStore,diffDocuments} from '@wieslawsoltes/gridweb/collaboration';
import {createCollaborationServer,bearerAuthorization} from '@wieslawsoltes/gridweb/collaboration/server';
const book=new Workbook();
const session=new CollaborationSession(book,{url:'https://example.test',room:'team',token:'provided-at-runtime',storage:sessionStorage});
session.Changed.Subscribe(e=>console.log(e.status,e.pending));
session.PresenceChanged.Subscribe(p=>console.log(p.map(p=>p.selection)));
const delta=diffDocuments(book.ToJSON(),book.ToJSON());
const store=new CollaborationStore();
console.log(delta,store,createCollaborationServer,bearerAuthorization);
// @ts-expect-error a workbook is required
new CollaborationSession({},{});
// @ts-expect-error explicit supported strategy
session.Resolve('merge-everything');
