import { randomUUID } from "node:crypto";
import { Hono } from "hono";
import { bodyLimit } from "hono/body-limit";
import { serveStatic } from "@hono/node-server/serve-static";
import { ADDIN_ORIGIN, MAX_BASE64_LENGTH, RUN_TIMEOUT_MS, runRequestSchema, type RunRequest, type RunResponse } from "../contract.js";

type Runner = (request: RunRequest, signal: AbortSignal) => Promise<RunResponse>;

export function createApp(run: Runner | undefined) {
  const app = new Hono();
  let activeRuns = 0;
  app.use("*", async (c, next) => {
    if (c.req.header("host") !== new URL(ADDIN_ORIGIN).host) return c.json({ error: "Invalid host." }, 403);
    c.header("Cache-Control", "no-store");
    c.header("X-Content-Type-Options", "nosniff");
    c.header("Referrer-Policy", "no-referrer");
    await next();
  });
  app.get("/api/health", (c) => c.json({ ready: Boolean(run) }));
  app.use("/api/run", bodyLimit({ maxSize: MAX_BASE64_LENGTH + 200_000, onError: (c) => c.json({ error: "Document or conversation is too large." }, 413) }));
  app.post("/api/run", async (c) => {
    if (c.req.header("origin") !== ADDIN_ORIGIN || !c.req.header("content-type")?.startsWith("application/json")) {
      return c.json({ error: "Use the SuperDoc Word task pane." }, 403);
    }
    if (!run) return c.json({ error: "Set OPENAI_API_KEY in the backend .env file and restart the server." }, 503);
    if (activeRuns >= 2) return c.json({ error: "The agent is busy. Try again shortly." }, 429);
    const parsed = runRequestSchema.safeParse(await c.req.json().catch(() => null));
    if (!parsed.success) return c.json({ error: "Invalid document, prompt, or conversation." }, 400);
    activeRuns += 1;
    const requestId = randomUUID();
    try {
      const signal = AbortSignal.any([c.req.raw.signal, AbortSignal.timeout(RUN_TIMEOUT_MS)]);
      const result = await run(parsed.data, signal);
      signal.throwIfAborted();
      return c.json(result);
    } catch {
      console.error(JSON.stringify({ event: "word_agent_run_failed", requestId }));
      return c.json({ error: `The agent could not finish. Word has not been updated. Reference: ${requestId}` }, 502);
    } finally {
      activeRuns -= 1;
    }
  });
  app.get("/", (c) => c.redirect("/taskpane.html"));
  app.use("*", serveStatic({ root: "./dist/ui" }));
  return app;
}
