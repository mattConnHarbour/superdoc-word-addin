import { z } from "zod";

export const ADDIN_ORIGIN = "https://localhost:3000";
export const MAX_DOCX_BYTES = 32 * 1024 * 1024;
export const MAX_BASE64_LENGTH = 4 * Math.ceil(MAX_DOCX_BYTES / 3);
export const RUN_TIMEOUT_MS = 120_000;
export const chatMessageSchema = z.object({
  role: z.enum(["user", "assistant"]),
  content: z.string().min(1).max(12_000),
});
export type ChatMessage = z.infer<typeof chatMessageSchema>;
export const runRequestSchema = z.object({
  requestId: z.uuid(),
  prompt: z.string().trim().min(1).max(12_000),
  documentBase64: z.string().min(4).max(MAX_BASE64_LENGTH).regex(/^[A-Za-z0-9+/]+={0,2}$/u),
  history: z.array(chatMessageSchema).max(12),
});
export type RunRequest = z.infer<typeof runRequestSchema>;
export const fragmentPatchSchema = z.object({
  start: z.number().int().nonnegative(),
  deleteCount: z.number().int().nonnegative(),
  expected: z.array(z.string().regex(/^[a-f0-9]{64}$/u)),
  ooxml: z.string().max(2 * 1024 * 1024),
  insertCount: z.number().int().nonnegative(),
});
export type FragmentPatch = z.infer<typeof fragmentPatchSchema>;
export const runResponseSchema = z.object({
  requestId: z.uuid(),
  sourceDigest: z.string().regex(/^[a-f0-9]{64}$/u),
  message: z.string(),
  documentBase64: z.string().min(4).max(MAX_BASE64_LENGTH),
  toolsUsed: z.array(z.string()),
  patches: z.array(fragmentPatchSchema).max(10_000),
  unsupportedReason: z.string().optional(),
});
export type RunResponse = z.infer<typeof runResponseSchema>;
export type FragmentUpdate = Pick<RunResponse, "patches" | "unsupportedReason">;
export type FullDocumentUpdate = Pick<RunResponse, "documentBase64">;
export const MAX_FULL_REPLACEMENT_BYTES = 4 * 1024 * 1024;

export function errorMessage(error: unknown): string {
  return error instanceof Error ? error.message : "Something went wrong.";
}
