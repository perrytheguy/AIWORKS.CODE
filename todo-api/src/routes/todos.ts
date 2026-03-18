import { Hono } from "hono";
import db from "../db";

const todos = new Hono();

// GET /todos
todos.get("/", (c) => {
  const rows = db.query("SELECT * FROM todos ORDER BY createdAt DESC").all();
  return c.json(rows);
});

// POST /todos
todos.post("/", async (c) => {
  const { title, dueDate } = await c.req.json<{ title: string; dueDate?: string }>();
  if (!title?.trim()) {
    return c.json({ error: "title은 필수입니다." }, 400);
  }
  const result = db
    .query("INSERT INTO todos (title, dueDate) VALUES (?, ?) RETURNING *")
    .get(title.trim(), dueDate ?? null);
  return c.json(result, 201);
});

// PUT /todos/:id — 완료 처리 (toggle)
todos.put("/:id", (c) => {
  const id = Number(c.req.param("id"));
  const existing = db.query("SELECT * FROM todos WHERE id = ?").get(id) as any;
  if (!existing) return c.json({ error: "할일을 찾을 수 없습니다." }, 404);

  const updated = db
    .query("UPDATE todos SET completed = ? WHERE id = ? RETURNING *")
    .get(existing.completed ? 0 : 1, id);
  return c.json(updated);
});

// DELETE /todos/:id
todos.delete("/:id", (c) => {
  const id = Number(c.req.param("id"));
  const existing = db.query("SELECT id FROM todos WHERE id = ?").get(id);
  if (!existing) return c.json({ error: "할일을 찾을 수 없습니다." }, 404);

  db.query("DELETE FROM todos WHERE id = ?").run(id);
  return c.json({ message: "삭제되었습니다." });
});

export default todos;
