import { readFile } from "node:fs/promises";
import { createServer } from "node:https";
import { homedir } from "node:os";
import path from "node:path";
import { getRequestListener } from "@hono/node-server";
import { ADDIN_ORIGIN } from "../contract.js";
import { createCompletion, runAgent } from "./agent.js";
import { createApp } from "./app.js";

const certDirectory = path.join(homedir(), ".office-addin-dev-certs");
const [cert, key] = await Promise.all([
  readFile(path.join(certDirectory, "localhost.crt")),
  readFile(path.join(certDirectory, "localhost.key")),
]).catch(() => { throw new Error("Run pnpm certs to install the trusted localhost certificate before starting Word Agent."); });
const apiKey = process.env.OPENAI_API_KEY;
const complete = apiKey ? createCompletion(apiKey) : undefined;
const app = createApp(complete ? (request, signal) => runAgent(request, complete, signal) : undefined);
createServer({ cert, key }, getRequestListener(app.fetch)).listen(3000, "127.0.0.1", () => {
  console.log(`Word Agent: ${ADDIN_ORIGIN}/taskpane.html`);
  console.log(complete ? "Agent configured." : "Set OPENAI_API_KEY in .env to enable the agent.");
});
