export async function officeReady(): Promise<boolean> {
  if (typeof Office === "undefined") return false;
  const info = await Office.onReady();
  return info.host === Office.HostType.Word
    && Office.context.requirements.isSetSupported("WordApi", "1.5")
    && Office.context.requirements.isSetSupported("CompressedFile", "1.1");
}
