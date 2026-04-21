import { useState } from "react";
import { Button } from "@/shared/ui/Button";
import { Alert } from "@/shared/ui/Alert";
import { archiveApi } from "@/features/commercial-archive/api/archiveApi";
import { getErrorMessage } from "@/shared/lib/apiError";

export const CurrentPlanButton = () => {
  const [error, setError] = useState<string | null>(null);
  const [isLoading, setLoading] = useState(false);

  const onClick = async () => {
    setError(null);
    setLoading(true);
    try {
      const response = await fetch(archiveApi.buildCurrentPlanUrl(), {
        method: "GET",
        credentials: "include",
      });
      if (!response.ok) {
        let detail = response.statusText || "Не удалось собрать диаграмму Ганта.";
        try {
          const data = (await response.json()) as { detail?: string };
          if (data.detail) {
            detail = data.detail;
          }
        } catch {
          // keep statusText
        }
        throw new Error(detail);
      }
      const blob = await response.blob();
      const url = window.URL.createObjectURL(blob);
      const anchor = document.createElement("a");
      anchor.href = url;
      anchor.download = `Актуальный_план_${new Date().toISOString().slice(0, 10)}.xlsx`;
      document.body.appendChild(anchor);
      anchor.click();
      anchor.remove();
      window.URL.revokeObjectURL(url);
    } catch (err) {
      setError(getErrorMessage(err));
    } finally {
      setLoading(false);
    }
  };

  return (
    <div style={{ display: "grid", gap: "0.5rem" }}>
      <Button variant="secondary" onClick={onClick} disabled={isLoading}>
        {isLoading ? "Собираю диаграмму..." : "📊 Актуальный план"}
      </Button>
      {error && <Alert tone="error">{error}</Alert>}
    </div>
  );
};
