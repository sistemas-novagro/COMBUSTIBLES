import { MongoClient } from 'mongodb';

const defaultData = { combustible: [], saldosReferencia: {} };

const client = new MongoClient(process.env.MONGODB_URI);
await client.connect();
const coleccion = client.db('novagro').collection('estado');

// Este objeto imita exactamente la forma en que usábamos lowdb antes
// (db.data.combustible, db.data.saldosReferencia, await db.read(), await db.write())
// así que el resto de server.js no necesita cambiar nada.
const db = {
  data: structuredClone(defaultData),
  async read() {
    const doc = await coleccion.findOne({ _id: 'principal' });
    if (doc) {
      delete doc._id;
      db.data = doc;
    } else {
      db.data = structuredClone(defaultData);
    }
    db.data.combustible ||= [];
    db.data.saldosReferencia ||= {};
  },
  async write() {
    await coleccion.updateOne(
      { _id: 'principal' },
      { $set: db.data },
      { upsert: true }
    );
  },
};

await db.read();
await db.write();

export default db;