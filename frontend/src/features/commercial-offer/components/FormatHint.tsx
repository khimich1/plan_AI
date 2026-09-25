import { useEffect, useRef, useState } from "react";

import { FORMAT_HINT_SHARED } from "@/features/commercial-offer/lib/productTypeConfig";

const COPY_SUCCESS_RESET_MS = 2000;

const toggleButtonStyle = {
  border: "none",
  background: "none",
  color: "#175cd3",
  cursor: "pointer",
  padding: "0.25rem 0",
  font: "inherit",
  textAlign: "left" as const,
};

type FormatHintProps = {
  explanation: string;
  sample: string;
};

type CopyStatus = "idle" | "copied" | "failed";

export const FormatHint = ({ explanation, sample }: FormatHintProps) => {
  const [open, setOpen] = useState(false);
  const [copyStatus, setCopyStatus] = useState<CopyStatus>("idle");
  const resetTimerRef = useRef<ReturnType<typeof setTimeout> | null>(null);

  useEffect(() => {
    return () => {
      if (resetTimerRef.current != null) {
        clearTimeout(resetTimerRef.current);
      }
    };
  }, []);

  const copySample = async () => {
    if (resetTimerRef.current != null) {
      clearTimeout(resetTimerRef.current);
      resetTimerRef.current = null;
    }
    try {
      await navigator.clipboard.writeText(sample);
      setCopyStatus("copied");
      resetTimerRef.current = setTimeout(() => {
        setCopyStatus("idle");
        resetTimerRef.current = null;
      }, COPY_SUCCESS_RESET_MS);
    } catch {
      setCopyStatus("failed");
    }
  };

  const copyButtonLabel =
    copyStatus === "copied" ? "Скопировано" : "Скопировать образец";

  return (
    <div>
      <button
        type="button"
        aria-expanded={open}
        onClick={() => setOpen((value) => !value)}
        style={toggleButtonStyle}
      >
        {open ? "▾ Подсказка" : "▸ Подсказка"}
      </button>
      {open && (
        <div style={{ display: "grid", gap: "0.5rem", marginTop: "0.35rem" }}>
          <p style={{ margin: 0, color: "#344054", fontSize: "0.9rem" }}>{FORMAT_HINT_SHARED}</p>
          <p style={{ margin: 0, color: "#344054", fontSize: "0.9rem" }}>{explanation}</p>
          <pre
            style={{
              margin: 0,
              padding: "0.6rem 0.75rem",
              background: "#f9fafb",
              border: "1px solid #e4e7ec",
              borderRadius: 8,
              whiteSpace: "pre-wrap",
              fontFamily: "Consolas, monospace",
              fontSize: "0.9rem",
            }}
          >
            {sample}
          </pre>
          <div style={{ display: "flex", flexWrap: "wrap", gap: "0.75rem", alignItems: "center" }}>
            <button type="button" onClick={copySample} style={toggleButtonStyle}>
              {copyButtonLabel}
            </button>
            {copyStatus === "failed" && (
              <span style={{ color: "#b42318", fontSize: "0.9rem" }}>Не удалось скопировать</span>
            )}
          </div>
        </div>
      )}
    </div>
  );
};
