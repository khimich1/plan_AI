import type { PropsWithChildren } from "react";

type AlertProps = PropsWithChildren<{
  tone?: "info" | "warning" | "error" | "success";
}>;

const toneStyles = {
  info: { background: "#eef4ff", borderColor: "#bfd4ff", color: "#1d4ed8" },
  warning: { background: "#fffaeb", borderColor: "#fedf89", color: "#b54708" },
  error: { background: "#fef3f2", borderColor: "#fecdca", color: "#b42318" },
  success: { background: "#ecfdf3", borderColor: "#abefc6", color: "#067647" },
};

export const Alert = ({ tone = "info", children }: AlertProps) => (
  <div
    style={{
      border: `1px solid ${toneStyles[tone].borderColor}`,
      background: toneStyles[tone].background,
      color: toneStyles[tone].color,
      borderRadius: 14,
      padding: "0.9rem 1rem",
    }}
  >
    {children}
  </div>
);
