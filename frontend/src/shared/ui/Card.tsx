import type { PropsWithChildren, ReactNode } from "react";

type CardProps = PropsWithChildren<{
  title?: string;
  subtitle?: ReactNode;
  actions?: ReactNode;
}>;

export const Card = ({ title, subtitle, actions, children }: CardProps) => (
  <section
    style={{
      background: "#ffffff",
      borderRadius: 20,
      border: "1px solid #e4e7ec",
      padding: "1.25rem",
      boxShadow: "0 10px 30px rgba(15, 23, 42, 0.06)",
    }}
  >
    {(title || subtitle || actions) && (
      <div
        style={{
          display: "flex",
          justifyContent: "space-between",
          gap: "1rem",
          alignItems: "flex-start",
          marginBottom: "1rem",
        }}
      >
        <div>
          {title && <h2 style={{ margin: 0, fontSize: "1.125rem" }}>{title}</h2>}
          {subtitle && <div style={{ color: "#475467", marginTop: "0.35rem" }}>{subtitle}</div>}
        </div>
        {actions}
      </div>
    )}
    {children}
  </section>
);
