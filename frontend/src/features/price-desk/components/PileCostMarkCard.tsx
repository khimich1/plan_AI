import { Input } from "@/shared/ui/Field";
import type {
  PileCostGradeBreakdown,
  PileCostMarkBreakdown,
  PileCostMarkSummary,
} from "@/features/price-desk/types/pileCost";
import { articleLabel, formatPileAmount, gradeLabel } from "@/features/price-desk/components/pileCostFormat";

type PileCostMarkCardProps = {
  marks: PileCostMarkSummary[];
  query: string;
  selectedMark: string | null;
  breakdown: PileCostMarkBreakdown | null;
  loading: boolean;
  onQuery: (value: string) => void;
  onSelect: (mark: string) => void;
};

const matchesOf = (marks: PileCostMarkSummary[], query: string): PileCostMarkSummary[] => {
  const needle = query.trim().toLowerCase();
  if (!needle) {
    return [];
  }
  return marks.filter((item) => item.mark.toLowerCase().includes(needle)).slice(0, 20);
};

const KgLine = ({ label, kg, rub }: { label: string; kg: number; rub: number }) => (
  <li>
    {label}: {formatPileAmount(kg)} кг, {formatPileAmount(rub)} руб
  </li>
);

const GradeBlock = ({ grade }: { grade: PileCostGradeBreakdown }) => (
  <article
    aria-label={gradeLabel(grade.concrete_grade)}
    style={{
      border: "1px solid #e4e7ec",
      borderRadius: 12,
      padding: "0.75rem 0.85rem",
      display: "grid",
      gap: "0.35rem",
    }}
  >
    <h4 style={{ margin: 0, fontSize: "0.98rem" }}>{gradeLabel(grade.concrete_grade)}</h4>
    <p style={{ margin: 0, fontWeight: 700 }}>
      Цена покупателя: {formatPileAmount(grade.buyer_price)}
      {grade.surcharge_applied ? " · ×1,02" : ""}
    </p>
    <p style={{ margin: 0, color: "#344054" }}>
      Материалы {formatPileAmount(grade.materials_rub)} · расходы {formatPileAmount(grade.expenses_rub)} ·
      себестоимость {formatPileAmount(grade.cost)}
    </p>
    <ul style={{ margin: 0, paddingLeft: "1.1rem", color: "#344054" }}>
      <KgLine label="Цемент" kg={grade.cement_kg} rub={grade.cement_rub} />
      <KgLine label="Песок" kg={grade.sand_kg} rub={grade.sand_rub} />
      <KgLine
        label={`Щебень (${articleLabel(grade.crushed_article)})`}
        kg={grade.crushed_kg}
        rub={grade.crushed_rub}
      />
      <KgLine label="Добавка" kg={grade.additive_kg} rub={grade.additive_rub} />
      {grade.lines.map((line) => (
        <KgLine key={line.article} label={articleLabel(line.article)} kg={line.kg} rub={line.rub} />
      ))}
    </ul>
  </article>
);

export const PileCostMarkCard = ({
  marks,
  query,
  selectedMark,
  breakdown,
  loading,
  onQuery,
  onSelect,
}: PileCostMarkCardProps) => {
  const matches = matchesOf(marks, query);
  const showList = query.trim().length > 0 && query.trim() !== selectedMark;

  return (
    <section aria-label="Карточка марки" style={{ display: "grid", gap: "0.65rem" }}>
      <h3 style={{ margin: 0, fontSize: "1rem" }}>Марка</h3>
      <label style={{ display: "grid", gap: "0.35rem" }}>
        <span style={{ fontWeight: 600 }}>Найти марку</span>
        <Input
          aria-label="Найти марку"
          value={query}
          autoComplete="off"
          onChange={(event) => onQuery(event.target.value)}
          onKeyDown={(event) => {
            if (event.key === "Enter" && matches[0]) {
              event.preventDefault();
              onSelect(matches[0].mark);
            }
          }}
        />
      </label>
      {showList && matches.length === 0 && <p style={{ margin: 0, color: "#475467" }}>Марка не найдена</p>}
      {showList && matches.length > 0 && (
        <ul role="listbox" aria-label="Марки" style={{ listStyle: "none", margin: 0, padding: 0 }}>
          {matches.map((item) => (
            <li key={item.mark}>
              <button
                type="button"
                onClick={() => onSelect(item.mark)}
                style={{
                  width: "100%",
                  textAlign: "left",
                  background: "#fff",
                  border: "1px solid #e4e7ec",
                  borderRadius: 8,
                  padding: "0.45rem 0.6rem",
                  marginBottom: "0.3rem",
                  cursor: "pointer",
                }}
              >
                {item.mark}
              </button>
            </li>
          ))}
        </ul>
      )}
      {!selectedMark && (
        <p style={{ margin: 0, color: "#475467" }}>Найдите марку — откроется карточка пяти классов.</p>
      )}
      {selectedMark && loading && !breakdown && (
        <p style={{ margin: 0, color: "#475467" }}>Считаю марку…</p>
      )}
      {breakdown && (
        <div
          role="region"
          aria-label={`Разложение ${breakdown.mark}`}
          style={{ display: "grid", gap: "0.65rem" }}
        >
          <p style={{ margin: 0 }}>
            {breakdown.mark}, сечение {breakdown.section_cm} см. Объём {formatPileAmount(breakdown.volume_m3)} м³.
            Объём и килограммы арматуры только для просмотра.
          </p>
          {breakdown.grades.map((grade) => (
            <GradeBlock key={grade.concrete_grade} grade={grade} />
          ))}
        </div>
      )}
    </section>
  );
};
