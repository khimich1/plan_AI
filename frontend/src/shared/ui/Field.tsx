import type { CSSProperties, InputHTMLAttributes, PropsWithChildren, TextareaHTMLAttributes } from "react";

type FieldWrapperProps = PropsWithChildren<{
  label: string;
  error?: string;
  hint?: string;
}>;

export const FieldWrapper = ({ label, error, hint, children }: FieldWrapperProps) => (
  <label style={{ display: "grid", gap: "0.45rem" }}>
    <span style={{ fontWeight: 600 }}>{label}</span>
    {children}
    {hint && !error && <span style={{ color: "#667085", fontSize: "0.9rem" }}>{hint}</span>}
    {error && <span style={{ color: "#b42318", fontSize: "0.9rem" }}>{error}</span>}
  </label>
);

const inputStyle: CSSProperties = {
  width: "100%",
  border: "1px solid #d0d5dd",
  borderRadius: 12,
  padding: "0.8rem 0.9rem",
  background: "#ffffff",
};

export const Input = (props: InputHTMLAttributes<HTMLInputElement>) => <input {...props} style={inputStyle} />;

export const Textarea = (props: TextareaHTMLAttributes<HTMLTextAreaElement>) => (
  <textarea {...props} style={{ ...inputStyle, resize: "vertical", minHeight: 160 }} />
);
