import { Low } from 'lowdb';
import { JSONFile } from 'lowdb/node';
import path from 'path';

// En tu compu guarda data.json en esta misma carpeta (comportamiento de siempre).
// En Railway, DATA_DIR va a apuntar a la carpeta del disco persistente (ej. /data),
// así los datos no se pierden cuando el servicio reinicia.
const carpetaDatos = process.env.DATA_DIR || '.';
const adapter = new JSONFile(path.join(carpetaDatos, 'data.json'));

const defaultData = { combustible: [], saldosReferencia: {} };
const db = new Low(adapter, defaultData);

await db.read();
db.data ||= defaultData;
db.data.saldosReferencia ||= {};
await db.write();

export default db;