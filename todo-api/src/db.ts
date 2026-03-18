import { Database } from "bun:sqlite";

const db = new Database("todos.db", { create: true });

db.exec(`
  CREATE TABLE IF NOT EXISTS todos (
    id        INTEGER PRIMARY KEY AUTOINCREMENT,
    title     TEXT    NOT NULL,
    completed INTEGER NOT NULL DEFAULT 0,
    dueDate   TEXT,
    createdAt TEXT    NOT NULL DEFAULT (datetime('now', 'localtime'))
  )
`);

export default db;
