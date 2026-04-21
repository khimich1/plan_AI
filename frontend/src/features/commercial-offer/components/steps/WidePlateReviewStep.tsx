import type { CommercialDraftDetails, WidePlateAction } from "@/features/commercial-offer/types/commercialOffer";
import { Alert } from "@/shared/ui/Alert";
import { Button } from "@/shared/ui/Button";
import { Card } from "@/shared/ui/Card";
import { FieldWrapper, Textarea } from "@/shared/ui/Field";
import { StepLayout } from "@/shared/ui/StepLayout";

type WidePlateReviewStepProps = {
  draft: CommercialDraftDetails;
  decisions: Record<string, { action: WidePlateAction; replacementText: string }>;
  errorMessage: string | null;
  isPending: boolean;
  onDecisionChange: (line: string, action: WidePlateAction, replacementText: string) => void;
  onBack: () => void;
  onSubmit: () => void;
};

export const WidePlateReviewStep = ({
  draft,
  decisions,
  errorMessage,
  isPending,
  onDecisionChange,
  onBack,
  onSubmit,
}: WidePlateReviewStepProps) => (
  <StepLayout
    title="Шаг 2. Проверка проблемных плит"
    description="Backend обнаружил плиты шире 12 дм. Для каждой позиции выберите действие: подтвердить, заменить или исключить."
    footer={
      <div style={{ display: "flex", justifyContent: "space-between", gap: "0.75rem" }}>
        <Button type="button" variant="ghost" onClick={onBack}>
          Назад
        </Button>
        <Button type="button" onClick={onSubmit} disabled={isPending}>
          {isPending ? "Сохраняем..." : "Подтвердить обработку"}
        </Button>
      </div>
    }
  >
    {errorMessage && <Alert tone="error">{errorMessage}</Alert>}

    {draft.metadata.wide_plate_lines.length === 0 ? (
      <Alert tone="success">Проблемных плит нет. Можно перейти дальше.</Alert>
    ) : (
      draft.metadata.wide_plate_lines.map((item) => {
        const currentDecision = decisions[item.line] ?? { action: "confirm", replacementText: "" };

        return (
          <Card key={item.line} title={item.line} subtitle={`Количество: ${item.qty}`}>
            <div style={{ display: "grid", gap: "0.75rem" }}>
              <div style={{ display: "flex", gap: "0.75rem", flexWrap: "wrap" }}>
                {([
                  ["confirm", "Подтвердить"],
                  ["replace", "Заменить"],
                  ["exclude", "Исключить"],
                ] as Array<[WidePlateAction, string]>).map(([value, label]) => (
                  <label
                    key={value}
                    style={{
                      display: "flex",
                      gap: "0.5rem",
                      alignItems: "center",
                      border: "1px solid #e4e7ec",
                      borderRadius: 12,
                      padding: "0.7rem 0.9rem",
                    }}
                  >
                    <input
                      type="radio"
                      checked={currentDecision.action === value}
                      onChange={() => onDecisionChange(item.line, value, currentDecision.replacementText)}
                    />
                    <span>{label}</span>
                  </label>
                ))}
              </div>

              {currentDecision.action === "replace" && (
                <FieldWrapper
                  label="Текст замены"
                  hint="Можно указать одну или несколько заменяющих позиций. Каждая строка будет нормализована backend."
                >
                  <Textarea
                    value={currentDecision.replacementText}
                    onChange={(event) => onDecisionChange(item.line, "replace", event.target.value)}
                    placeholder={"ПБ 78-12-8п 2\nПБ 78-3-8п 2"}
                  />
                </FieldWrapper>
              )}
            </div>
          </Card>
        );
      })
    )}
  </StepLayout>
);
