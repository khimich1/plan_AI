import { useState } from "react";
import {
  GOST_FROST_RESISTANCE,
  GOST_WATERPROOFNESS,
  formatFrostPair,
  isOffTablePair,
  tablePairsForGrade,
  type ConcreteSpecFields,
  type ConcreteSpecPatch,
} from "@/features/commercial-offer/lib/concreteSpec";

type ConcreteSpecCellProps = {
  grade: string;
  mark: string;
  spec: ConcreteSpecFields;
  sealed: boolean;
  onChange?: (patch: ConcreteSpecPatch) => void;
};

const selectStyle = {
  border: "1px solid #d0d5dd",
  borderRadius: 8,
  padding: "0.35rem 0.5rem",
  background: "#ffffff",
};

export const ConcreteSpecCell = ({ grade, mark, spec, sealed, onChange }: ConcreteSpecCellProps) => {
  const label = `F / W ${mark}`;
  const text = formatFrostPair(spec.frost_resistance, spec.waterproofness);
  const [otherOpen, setOtherOpen] = useState(false);
  const pairs = tablePairsForGrade(grade);
  const offTable = isOffTablePair(grade, spec.frost_resistance, spec.waterproofness);
  const matched =
    pairs.find(
      (item) =>
        item.frost_resistance === spec.frost_resistance &&
        item.waterproofness === spec.waterproofness &&
        item.aggregate === spec.concrete_aggregate,
    ) ??
    pairs.find(
      (item) =>
        item.frost_resistance === spec.frost_resistance && item.waterproofness === spec.waterproofness,
    );
  const editable = Boolean(onChange) && !sealed && text !== "—";

  if (!editable) {
    return <span aria-label={label}>{text}</span>;
  }

  const selectValue = otherOpen ? "other" : offTable ? "saved" : (matched?.aggregate ?? "saved");
  const frostValue = GOST_FROST_RESISTANCE.includes(spec.frost_resistance as (typeof GOST_FROST_RESISTANCE)[number])
    ? spec.frost_resistance
    : GOST_FROST_RESISTANCE[0];
  const waterValue = GOST_WATERPROOFNESS.includes(spec.waterproofness as (typeof GOST_WATERPROOFNESS)[number])
    ? spec.waterproofness
    : GOST_WATERPROOFNESS[0];

  return (
    <div style={{ display: "grid", gap: "0.35rem" }}>
      <select
        aria-label={label}
        value={selectValue}
        style={selectStyle}
        onChange={(event) => {
          const next = event.target.value;
          if (next === "other") {
            setOtherOpen(true);
            return;
          }
          setOtherOpen(false);
          if (next === "granite" || next === "ordinary") {
            onChange?.({ concrete_spec_source: "table", concrete_aggregate: next });
          }
        }}
      >
        {offTable ? <option value="saved">{text}</option> : null}
        {pairs.map((item) => (
          <option key={item.aggregate} value={item.aggregate}>
            {item.label}
          </option>
        ))}
        <option value="other">Другое…</option>
      </select>
      {offTable && !otherOpen ? (
        <span style={{ color: "#b54708", fontSize: "0.8rem" }}>не по таблице</span>
      ) : null}
      {otherOpen ? (
        <div style={{ display: "flex", gap: "0.35rem", flexWrap: "wrap" }}>
          <select
            aria-label={`Морозостойкость ${mark}`}
            value={frostValue ?? ""}
            style={selectStyle}
            onChange={(event) =>
              onChange?.({
                concrete_spec_source: "manual",
                frost_resistance: event.target.value,
                waterproofness: waterValue ?? "",
              })
            }
          >
            {GOST_FROST_RESISTANCE.map((code) => (
              <option key={code} value={code}>
                {code}
              </option>
            ))}
          </select>
          <select
            aria-label={`Водонепроницаемость ${mark}`}
            value={waterValue ?? ""}
            style={selectStyle}
            onChange={(event) =>
              onChange?.({
                concrete_spec_source: "manual",
                frost_resistance: frostValue ?? "",
                waterproofness: event.target.value,
              })
            }
          >
            {GOST_WATERPROOFNESS.map((code) => (
              <option key={code} value={code}>
                {code}
              </option>
            ))}
          </select>
        </div>
      ) : null}
    </div>
  );
};
