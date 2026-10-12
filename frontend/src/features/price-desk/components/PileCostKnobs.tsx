import { FieldWrapper, Input } from "@/shared/ui/Field";
import type { PileCostInputsBody } from "@/features/price-desk/types/pileCost";
import {
  articleLabel,
  gradeLabel,
  type KnobDraft,
  type OverheadField,
  type RecipeField,
} from "@/features/price-desk/components/pileCostFormat";

const OVERHEAD_FIELDS: { field: OverheadField; label: string }[] = [
  { field: "general", label: "Общие производственные" },
  { field: "payroll", label: "ФОТ" },
  { field: "forming", label: "Обслуживание формовки" },
  { field: "other", label: "Обслуживание, остальное" },
];

const RECIPE_FIELDS: { field: RecipeField; label: string }[] = [
  { field: "cement_kg_m3", label: "Цемент" },
  { field: "sand_kg_m3", label: "Песок" },
  { field: "crushed_kg_m3", label: "Щебень" },
];

type PileCostKnobsProps = {
  source: PileCostInputsBody;
  draft: KnobDraft;
  onChange: (draft: KnobDraft) => void;
  onCommit: (draft: KnobDraft) => void;
};

export const PileCostKnobs = ({ source, draft, onChange, onCommit }: PileCostKnobsProps) => {
  const commitPrice = (article: string, value: string) => {
    const next: KnobDraft = {
      ...draft,
      prices: { ...draft.prices, [article]: value },
    };
    onCommit(next);
  };

  const commitRecipe = (grade: string, field: RecipeField, value: string) => {
    const current = draft.recipes[grade] ?? {
      cement_kg_m3: "",
      sand_kg_m3: "",
      crushed_kg_m3: "",
    };
    const next: KnobDraft = {
      ...draft,
      recipes: { ...draft.recipes, [grade]: { ...current, [field]: value } },
    };
    onCommit(next);
  };

  const commitOverhead = (field: OverheadField, value: string) => {
    const next: KnobDraft = {
      ...draft,
      overhead: { ...draft.overhead, [field]: value },
    };
    onCommit(next);
  };

  return (
    <div style={{ display: "grid", gap: "1rem" }}>
      <section aria-label="Цены статей" style={{ display: "grid", gap: "0.65rem" }}>
        <h3 style={{ margin: 0, fontSize: "1rem" }}>Цены</h3>
        <div
          style={{
            display: "grid",
            gridTemplateColumns: "repeat(auto-fill, minmax(160px, 1fr))",
            gap: "0.65rem",
          }}
        >
          {source.prices.map((row) => (
            <FieldWrapper key={row.article} label={articleLabel(row.article)}>
              <Input
                inputMode="decimal"
                autoComplete="off"
                aria-label={`Цена: ${articleLabel(row.article)}`}
                value={draft.prices[row.article] ?? ""}
                onChange={(event) =>
                  onChange({
                    ...draft,
                    prices: { ...draft.prices, [row.article]: event.target.value },
                  })
                }
                onBlur={(event) => commitPrice(row.article, event.target.value)}
              />
            </FieldWrapper>
          ))}
        </div>
      </section>

      <section aria-label="Рецепт бетона" style={{ display: "grid", gap: "0.65rem" }}>
        <h3 style={{ margin: 0, fontSize: "1rem" }}>Рецепт по классу, кг/м³</h3>
        <div style={{ overflowX: "auto" }}>
          <table style={{ width: "100%", borderCollapse: "collapse", fontSize: "0.95rem" }}>
            <thead>
              <tr style={{ textAlign: "left", color: "#475467" }}>
                <th style={{ padding: "0.4rem 0.45rem" }}>Класс</th>
                {RECIPE_FIELDS.map((column) => (
                  <th key={column.field} style={{ padding: "0.4rem 0.45rem" }}>
                    {column.label}
                  </th>
                ))}
              </tr>
            </thead>
            <tbody>
              {source.recipes.map((row) => (
                <tr key={row.concrete_grade}>
                  <th scope="row" style={{ padding: "0.35rem 0.45rem", fontWeight: 600 }}>
                    {gradeLabel(row.concrete_grade)}
                  </th>
                  {RECIPE_FIELDS.map((column) => (
                    <td key={column.field} style={{ padding: "0.35rem 0.45rem" }}>
                      <Input
                        inputMode="decimal"
                        autoComplete="off"
                        aria-label={`${column.label} рецепта ${gradeLabel(row.concrete_grade)}`}
                        value={draft.recipes[row.concrete_grade]?.[column.field] ?? ""}
                        onChange={(event) => {
                          const current = draft.recipes[row.concrete_grade] ?? {
                            cement_kg_m3: "",
                            sand_kg_m3: "",
                            crushed_kg_m3: "",
                          };
                          onChange({
                            ...draft,
                            recipes: {
                              ...draft.recipes,
                              [row.concrete_grade]: { ...current, [column.field]: event.target.value },
                            },
                          });
                        }}
                        onBlur={(event) =>
                          commitRecipe(row.concrete_grade, column.field, event.target.value)
                        }
                      />
                    </td>
                  ))}
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      </section>

      <section aria-label="Статьи расходов" style={{ display: "grid", gap: "0.65rem" }}>
        <h3 style={{ margin: 0, fontSize: "1rem" }}>Расходы, руб/м³</h3>
        <div
          style={{
            display: "grid",
            gridTemplateColumns: "repeat(auto-fill, minmax(180px, 1fr))",
            gap: "0.65rem",
          }}
        >
          {OVERHEAD_FIELDS.map((row) => (
            <FieldWrapper key={row.field} label={row.label}>
              <Input
                inputMode="decimal"
                autoComplete="off"
                aria-label={row.label}
                value={draft.overhead[row.field] ?? ""}
                onChange={(event) =>
                  onChange({
                    ...draft,
                    overhead: { ...draft.overhead, [row.field]: event.target.value },
                  })
                }
                onBlur={(event) => commitOverhead(row.field, event.target.value)}
              />
            </FieldWrapper>
          ))}
        </div>
      </section>
    </div>
  );
};
