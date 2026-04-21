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
  const url = URL.createObjectURL(blob);
  const anchor = document.createElement("a");
  anchor.href = url;
  anchor.download = filename;
  anchor.rel = "noopener";
  document.body.appendChild(anchor);
  anchor.click();
  anchor.remove();
  setTimeout(() => URL.revokeObjectURL(url), 0);
};
