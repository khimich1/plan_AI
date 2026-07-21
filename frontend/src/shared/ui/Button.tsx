import type { ButtonHTMLAttributes, CSSProperties, PropsWithChildren } from "react";

type ButtonProps = PropsWithChildren<
  ButtonHTMLAttributes<HTMLButtonElement> & {
    variant?: "primary" | "secondary" | "ghost" | "danger";
    fullWidth?: boolean;
  }
>;

const variantStyles: Record<NonNullable<ButtonProps["variant"]>, CSSProperties> = {
  primary: {
    background: "#2b5cff",
    color: "#ffffff",
    border: "1px solid #2b5cff",
  },
  secondary: {
    background: "#eef2ff",
    color: "#23366f",
    border: "1px solid #d6defa",
  },
  ghost: {
    background: "transparent",
    color: "#23366f",
    border: "1px solid #d6defa",
  },
  danger: {
    background: "#fff1f2",
    color: "#b42318",
    border: "1px solid #fecdd3",
  },
};

export const Button = ({
  children,
  variant = "primary",
  fullWidth = false,
  style,
  disabled,
  ...props
}: ButtonProps) => (
  <button
    {...props}
    disabled={disabled}
    style={{
      borderRadius: 12,
      padding: "0.75rem 1rem",
      cursor: disabled ? "not-allowed" : "pointer",
      fontWeight: 600,
      opacity: disabled ? 0.6 : 1,
      width: fullWidth ? "100%" : undefined,
      ...variantStyles[variant],
      ...style,
    }}
  >
    {children}
  </button>
);
