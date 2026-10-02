import { Low } from 'lowdb';
import { JSONFile } from 'lowdb/node';

const adapter = new JSONFile('data.json');
const defaultData = { combustible: [], saldosReferencia: {} };
const db = new Low(adapter, defaultData);

await db.read();
db.data ||= defaultData;
db.data.saldosReferencia ||= {}; // por si el archivo ya existía sin este campo
await db.write();

export default db;