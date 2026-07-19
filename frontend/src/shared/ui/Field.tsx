import { useCallback, useLayoutEffect, useRef, type CSSProperties, type InputHTMLAttributes, type PropsWithChildren, type TextareaHTMLAttributes } from "react";

type FieldWrapperProps = PropsWithChildren<{
  label: string;
  error?: string | null;
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

export const AutoResizeTextarea = ({ value, onChange, style, ...props }: TextareaHTMLAttributes<HTMLTextAreaElement>) => {
  const ref = useRef<HTMLTextAreaElement>(null);

  const adjustHeight = useCallback(() => {
    const element = ref.current;
    if (!element) {
      return;
    }
    element.style.height = "auto";
    element.style.height = `${element.scrollHeight}px`;
  }, []);

  useLayoutEffect(() => {
    adjustHeight();
  }, [value, adjustHeight]);

  return (
    <textarea
      ref={ref}
      value={value}
      onChange={(event) => {
        onChange?.(event);
        adjustHeight();
      }}
      {...props}
      style={{
        ...inputStyle,
        resize: "none",
        overflow: "hidden",
        minHeight: 0,
        fontFamily: "Consolas, monospace",
        ...style,
      }}
    />
  );
};
