import { env } from "@/shared/config/env";

const resolveDownloadUrl = (downloadUrl: string): string => {
  const trimmedUrl = downloadUrl.trim();
  if (!trimmedUrl) {
    return trimmedUrl;
  }
  if (/^https?:\/\//i.test(trimmedUrl)) {
    return trimmedUrl;
  }
  const normalizedPath = trimmedUrl.startsWith("/") ? trimmedUrl : `/${trimmedUrl}`;
  return `${env.apiBaseUrl}${normalizedPath}`;
};

export const downloadFile = (downloadUrl: string): void => {
  window.open(resolveDownloadUrl(downloadUrl), "_blank", "noopener,noreferrer");
};

export const saveBlobAs = (blob: Blob, filename: string): void => {
  const file =
    blob.type === "application/pdf"
      ? new Blob([blob], { type: "application/octet-stream" })
      : blob;
  const url = URL.createObjectURL(file);
  const anchor = document.createElement("a");
  anchor.href = url;
  anchor.download = filename;
  document.body.appendChild(anchor);
  anchor.click();
  anchor.remove();
  window.setTimeout(() => URL.revokeObjectURL(url), 60_000);
};
