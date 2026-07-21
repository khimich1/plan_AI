import type { PropsWithChildren, ReactNode } from "react";

type StepLayoutProps = PropsWithChildren<{
  title: string;
  description: string;
  footer?: ReactNode;
  aside?: ReactNode;
}>;

export const StepLayout = ({ title, description, footer, aside, children }: StepLayoutProps) => (
  <div
    style={{
      display: "grid",
      gridTemplateColumns: aside ? "minmax(0, 1fr) 280px" : "minmax(0, 1fr)",
      gap: "1.25rem",
      alignItems: "start",
    }}
  >
    <div style={{ display: "grid", gap: "1rem" }}>
      <header>
        <h1 style={{ margin: 0, fontSize: "1.75rem" }}>{title}</h1>
        <p style={{ margin: "0.5rem 0 0", color: "#475467" }}>{description}</p>
      </header>
      {children}
      {footer}
    </div>
    {aside}
  </div>
);
