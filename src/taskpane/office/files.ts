import { MAX_DOCX_BYTES } from "../../contract.js";

export function bytesToBase64(bytes: Uint8Array): string {
  let binary = "";
  for (let offset = 0; offset < bytes.length; offset += 32_768) {
    binary += String.fromCharCode(...bytes.subarray(offset, offset + 32_768));
  }
  return btoa(binary);
}

export function readCompressedDocument(host: Pick<Office.Document, "getFileAsync">): Promise<Uint8Array<ArrayBuffer>> {
  return new Promise((resolve, reject) => {
    host.getFileAsync(Office.FileType.Compressed, { sliceSize: 65_536 }, (result) => {
      if (result.status !== Office.AsyncResultStatus.Succeeded) {
        reject(new Error(result.error.message));
        return;
      }
      const file = result.value;
      void (async () => {
        try {
          if (file.size > MAX_DOCX_BYTES) throw new Error("Word Agent supports DOCX files up to 32 MB.");
          const bytes = new Uint8Array(file.size);
          let offset = 0;
          for (let index = 0; index < file.sliceCount; index += 1) {
            const data = await new Promise<Uint8Array<ArrayBuffer>>((resolveSlice, rejectSlice) => {
              file.getSliceAsync(index, (slice) => {
                if (slice.status !== Office.AsyncResultStatus.Succeeded) rejectSlice(new Error(slice.error.message));
                else if (Array.isArray(slice.value.data) && slice.value.data.every((value: unknown) => typeof value === "number")) {
                  resolveSlice(new Uint8Array(slice.value.data));
                } else rejectSlice(new Error("Word returned an unsupported file slice."));
              });
            });
            bytes.set(data, offset);
            offset += data.length;
          }
          if (offset !== file.size) throw new Error("Word returned an incomplete document.");
          resolve(bytes);
        } catch (error) {
          reject(error);
        } finally {
          file.closeAsync();
        }
      })();
    });
  });
}
