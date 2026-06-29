type Props = {
  step: 1 | 2 | 3;
};

export const WizardStepIndicator = ({ step }: Props) => (
  <div style={{ display: "flex", gap: "0.5rem", marginBottom: "1rem" }}>
    {[1, 2, 3].map((s) => (
      <div
        key={s}
        style={{
          flex: 1,
          padding: "0.5rem 0.75rem",
          borderRadius: 10,
          background: s === step ? "#2b5cff" : "#eef2ff",
          color: s === step ? "#ffffff" : "#23366f",
          fontWeight: 600,
          textAlign: "center",
        }}
      >
        Шаг {s}
      </div>
    ))}
  </div>
);
