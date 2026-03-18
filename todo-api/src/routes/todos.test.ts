import { describe, it, expect, beforeAll } from "bun:test";

const BASE = "http://localhost:3001/todos";

// Helper: create a todo and return its parsed JSON body
async function createTodo(title: string, dueDate?: string) {
  const res = await fetch(BASE, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ title, ...(dueDate ? { dueDate } : {}) }),
  });
  return { res, body: await res.json() as any };
}

// Helper: delete a todo by id (best-effort cleanup)
async function deleteTodo(id: number) {
  await fetch(`${BASE}/${id}`, { method: "DELETE" });
}

// ─────────────────────────────────────────────────────────────────────────────
// Connectivity check — all suites depend on the server being reachable
// ─────────────────────────────────────────────────────────────────────────────
beforeAll(async () => {
  try {
    const res = await fetch(BASE);
    if (!res.ok) throw new Error(`Unexpected status ${res.status}`);
  } catch (e) {
    throw new Error(
      `Todo API server is not reachable at ${BASE}. Start it with 'bun run dev' before running tests.\n${e}`
    );
  }
});

// ─────────────────────────────────────────────────────────────────────────────
// GET /todos
// ─────────────────────────────────────────────────────────────────────────────
describe("GET /todos", () => {
  it("returns 200 and an array", async () => {
    const res = await fetch(BASE);
    const body = await res.json() as any[];

    expect(res.status).toBe(200);
    expect(Array.isArray(body)).toBe(true);
  });

  it("each item has id, title, completed, createdAt fields", async () => {
    const { body: created } = await createTodo("[GET-list] field check");

    const res = await fetch(BASE);
    const body = await res.json() as any[];

    const found = body.find((t: any) => t.id === created.id);
    expect(found).toBeDefined();
    expect(typeof found.id).toBe("number");
    expect(typeof found.title).toBe("string");
    expect(found.completed === 0 || found.completed === 1).toBe(true);
    expect(typeof found.createdAt).toBe("string");

    await deleteTodo(created.id);
  });
});

// ─────────────────────────────────────────────────────────────────────────────
// POST /todos
// ─────────────────────────────────────────────────────────────────────────────
describe("POST /todos", () => {
  it("creates a todo and returns 201 with the new item", async () => {
    const { res, body } = await createTodo("[POST] create basic todo");

    expect(res.status).toBe(201);
    expect(typeof body.id).toBe("number");
    expect(body.title).toBe("[POST] create basic todo");
    expect(body.completed).toBe(0);

    await deleteTodo(body.id);
  });

  it("stores dueDate when provided", async () => {
    const { res, body } = await createTodo("[POST] with dueDate", "2026-12-31");

    expect(res.status).toBe(201);
    expect(body.dueDate).toBe("2026-12-31");

    await deleteTodo(body.id);
  });

  it("trims leading/trailing whitespace from title", async () => {
    const { res, body } = await createTodo("  [POST] trimmed  ");

    expect(res.status).toBe(201);
    expect(body.title).toBe("[POST] trimmed");

    await deleteTodo(body.id);
  });

  it("returns 400 when title is missing", async () => {
    const res = await fetch(BASE, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ dueDate: "2026-01-01" }),
    });
    const body = await res.json() as any;

    expect(res.status).toBe(400);
    expect(body.error).toBeDefined();
  });

  it("returns 400 when title is an empty string", async () => {
    const res = await fetch(BASE, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ title: "" }),
    });
    const body = await res.json() as any;

    expect(res.status).toBe(400);
    expect(body.error).toBeDefined();
  });

  it("returns 400 when title is only whitespace", async () => {
    const res = await fetch(BASE, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ title: "   " }),
    });
    const body = await res.json() as any;

    expect(res.status).toBe(400);
    expect(body.error).toBeDefined();
  });
});

// ─────────────────────────────────────────────────────────────────────────────
// PUT /todos/:id  (toggle completed)
// ─────────────────────────────────────────────────────────────────────────────
describe("PUT /todos/:id", () => {
  it("toggles completed from 0 to 1 and returns updated item", async () => {
    const { body: created } = await createTodo("[PUT] toggle on");
    expect(created.completed).toBe(0);

    const res = await fetch(`${BASE}/${created.id}`, { method: "PUT" });
    const body = await res.json() as any;

    expect(res.status).toBe(200);
    expect(body.id).toBe(created.id);
    expect(body.completed).toBe(1);

    await deleteTodo(created.id);
  });

  it("toggles completed back from 1 to 0 on second call", async () => {
    const { body: created } = await createTodo("[PUT] toggle off");

    // First toggle: 0 -> 1
    await fetch(`${BASE}/${created.id}`, { method: "PUT" });

    // Second toggle: 1 -> 0
    const res = await fetch(`${BASE}/${created.id}`, { method: "PUT" });
    const body = await res.json() as any;

    expect(res.status).toBe(200);
    expect(body.completed).toBe(0);

    await deleteTodo(created.id);
  });

  it("returns 404 for a non-existent id", async () => {
    const res = await fetch(`${BASE}/999999999`, { method: "PUT" });
    const body = await res.json() as any;

    expect(res.status).toBe(404);
    expect(body.error).toBeDefined();
  });

  it("returns 404 for id 0", async () => {
    const res = await fetch(`${BASE}/0`, { method: "PUT" });
    const body = await res.json() as any;

    expect(res.status).toBe(404);
    expect(body.error).toBeDefined();
  });
});

// ─────────────────────────────────────────────────────────────────────────────
// DELETE /todos/:id
// ─────────────────────────────────────────────────────────────────────────────
describe("DELETE /todos/:id", () => {
  it("deletes an existing todo and returns success message", async () => {
    const { body: created } = await createTodo("[DELETE] remove me");

    const res = await fetch(`${BASE}/${created.id}`, { method: "DELETE" });
    const body = await res.json() as any;

    expect(res.status).toBe(200);
    expect(body.message).toBeDefined();
  });

  it("deleted item is no longer present in GET /todos", async () => {
    const { body: created } = await createTodo("[DELETE] verify gone");

    await fetch(`${BASE}/${created.id}`, { method: "DELETE" });

    const listRes = await fetch(BASE);
    const list = await listRes.json() as any[];
    const found = list.find((t: any) => t.id === created.id);

    expect(found).toBeUndefined();
  });

  it("returns 404 when deleting a non-existent id", async () => {
    const res = await fetch(`${BASE}/999999999`, { method: "DELETE" });
    const body = await res.json() as any;

    expect(res.status).toBe(404);
    expect(body.error).toBeDefined();
  });

  it("returns 404 on double-delete of the same id", async () => {
    const { body: created } = await createTodo("[DELETE] double delete");

    await fetch(`${BASE}/${created.id}`, { method: "DELETE" });

    const res = await fetch(`${BASE}/${created.id}`, { method: "DELETE" });
    const body = await res.json() as any;

    expect(res.status).toBe(404);
    expect(body.error).toBeDefined();
  });
});
