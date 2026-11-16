import "./taskpane.css";
import { errorMessage, type ChatMessage, type RunResponse } from "../contract.js";
import { officeReady } from "./office/host.js";
import { updateDocument, WordPatchError } from "./roundtrip.js";

function element<T extends HTMLElement>(id: string, type: new () => T): T {
  const value = document.getElementById(id);
  if (!(value instanceof type)) throw new Error(`Missing UI element: ${id}`);
  return value;
}

const form = element("chat-form", HTMLFormElement);
const prompt = element("prompt", HTMLTextAreaElement);
const send = element("send", HTMLButtonElement);
const cancel = element("cancel", HTMLButtonElement);
const status = element("status", HTMLParagraphElement);
const messages = element("messages", HTMLElement);
const history: ChatMessage[] = [];
let active: AbortController | undefined;

function addMessage(role: string, content: string) {
  const item = document.createElement("article");
  item.className = `message ${role === "You" ? "user" : "assistant"}`;
  const label = document.createElement("strong");
  label.textContent = role;
  const text = document.createElement("p");
  text.textContent = content;
  item.append(label, text);
  messages.append(item);
  messages.scrollTop = messages.scrollHeight;
  return item;
}

function addDownload(item: HTMLElement, response: RunResponse) {
  const button = document.createElement("button");
  button.className = "ds-button";
  button.textContent = "Download agent version";
  button.addEventListener("click", () => {
    const bytes = Uint8Array.from(atob(response.documentBase64), (character) => character.charCodeAt(0));
    const url = URL.createObjectURL(new Blob([bytes], { type: "application/vnd.openxmlformats-officedocument.wordprocessingml.document" }));
    const link = document.createElement("a");
    link.href = url;
    link.download = "agent-version.docx";
    link.click();
    setTimeout(() => URL.revokeObjectURL(url), 1_000);
  });
  item.append(button);
  messages.scrollTop = messages.scrollHeight;
}

form.addEventListener("submit", (event) => {
  event.preventDefault();
  const input = prompt.value.trim();
  if (!input || active) return;
  const controller = new AbortController();
  active = controller;
  send.disabled = true;
  prompt.disabled = true;
  cancel.hidden = false;
  addMessage("You", input);
  void (async () => {
    try {
      const result = await updateDocument(input, history, controller.signal, (text) => {
        status.textContent = text;
        cancel.disabled = text === "Updating Word…";
      });
      if (result.status === "conflict") {
        const item = addMessage("SuperDoc", "An affected paragraph or table changed while I was working. The patches have not been applied. You can download it or send your request again against the latest document.");
        addDownload(item, result.response);
        status.textContent = "Document changed · Word was not updated";
      } else if (result.status === "unsupported") {
        const item = addMessage("SuperDoc", `${result.response.unsupportedReason} Word has not been updated. You can download the agent version.`);
        addDownload(item, result.response);
        status.textContent = "Edit needs an unsupported fragment type";
      } else {
        addMessage("SuperDoc", result.response.message);
        history.push({ role: "user", content: input }, { role: "assistant", content: result.response.message.slice(0, 12_000) });
        history.splice(0, Math.max(0, history.length - 12));
        status.textContent = result.status === "applied" ? `${result.response.patches.length} fragment patch(es) applied` : "No document changes";
        prompt.value = "";
      }
    } catch (error) {
      status.textContent = controller.signal.aborted ? "Stopped" : "Could not complete the update";
      const item = addMessage("SuperDoc", controller.signal.aborted ? "Stopped before applying an agent result. You can send another request." : errorMessage(error));
      if (error instanceof WordPatchError) addDownload(item, error.response);
    } finally {
      active = undefined;
      send.disabled = false;
      prompt.disabled = false;
      cancel.hidden = true;
      cancel.disabled = false;
      prompt.focus();
    }
  })();
});

cancel.addEventListener("click", () => {
  cancel.disabled = true;
  active?.abort();
  cancel.disabled = false;
});

void (async () => {
  try {
    if (!await officeReady()) {
      status.textContent = "Open this add-in in desktop Word with WordApi 1.5 and compressed file access.";
      return;
    }
    const health = await fetch("/api/health");
    const body: unknown = await health.json();
    if (!body || typeof body !== "object" || !("ready" in body) || body.ready !== true) {
      status.textContent = "Set OPENAI_API_KEY in the backend .env file and restart the server.";
      return;
    }
    send.disabled = false;
    prompt.disabled = false;
    status.textContent = "Connected to Word";
  } catch (error) {
    status.textContent = errorMessage(error);
  }
})();
