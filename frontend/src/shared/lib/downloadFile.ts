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
