import { Hono } from "hono";
import { logger } from "hono/logger";
import todos from "./routes/todos";

const app = new Hono();

app.use("*", logger());
app.route("/todos", todos);

export default {
  port: 3001,
  fetch: app.fetch,
};

console.log("Todo API 서버 실행 중: http://localhost:3001");
