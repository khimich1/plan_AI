import { useEffect, useRef, useState } from "react";
import { Alert } from "@/shared/ui/Alert";
import { Button } from "@/shared/ui/Button";
import { getErrorMessage } from "@/shared/lib/apiError";
import { useDebouncedValue } from "@/shared/lib/useDebouncedValue";
import {
  useGuidExplainQuery,
  useGuidSearchQuery,
} from "@/features/nomenclature/hooks/useGuidExplainQueries";
import {
  GUID_SEARCH_MIN_CHARS,
  type GuidCandidate,
  type GuidCardChoice,
  type GuidExplainSource,
  type GuidReplacePreview,
} from "@/features/nomenclature/types/nomenclature";

type GuidExplainPanelProps = {
  source: GuidExplainSource;
  productKind: string;
  mark: string;
  /** После успешной записи: обновить список задач. Колокольчик панель не гасит. */
  onCommitted?: () => void;
};

const listStyle = { margin: "0.25rem 0 0", paddingLeft: "1.15rem", display: "grid", gap: "0.2rem" };
const headingStyle = { margin: "0.6rem 0 0", fontSize: "0.9rem" };

const cardButtonLabel = (card: GuidCardChoice, cards: GuidCardChoice[]): string => {
  const name = card.name.trim();
  const tail = card.guid.slice(-8);
  const shared = cards.filter((item) => item.name.trim() === name).length > 1;
  if (!name) {
    return tail;
  }
  if (!shared) {
    return name;
  }
  return `${name} · ${tail}`;
};

const rub = (value: number) =>
  new Intl.NumberFormat("ru-RU", { minimumFractionDigits: 2, maximumFractionDigits: 2 }).format(
    value,
  );

const CandidateRow = ({
  item,
  onReplace,
}: {
  item: GuidCandidate;
  onReplace?: (item: GuidCandidate) => void;
}) => (
  <li>
    {item.mark} — номеров: {item.guid_count}
    {item.guid_count === 0 ? " (номера нет)" : ""}
    {item.label ? ` · ${item.label}` : ""}
    {onReplace && item.selectable ? (
      <Button type="button" style={{ marginLeft: "0.5rem" }} onClick={() => onReplace(item)}>
        Заменить
      </Button>
    ) : null}
  </li>
);

const CandidateList = ({
  items,
  testId,
  onReplace,
}: {
  items: GuidCandidate[];
  testId: string;
  onReplace?: (item: GuidCandidate) => void;
}) => (
  <ul data-testid={testId} style={listStyle}>
    {items.map((item) => (
      <CandidateRow key={item.mark} item={item} onReplace={onReplace} />
    ))}
  </ul>
);

const GuidSearchBox = ({
  source,
  productKind,
  onReplace,
}: {
  source: GuidExplainSource;
  productKind: string;
  onReplace?: (item: GuidCandidate) => void;
}) => {
  const [text, setText] = useState("");
  const debounced = useDebouncedValue(text, 300);
  const search = useGuidSearchQuery(source, productKind, debounced);
  const tooShort = text.trim().length < GUID_SEARCH_MIN_CHARS;
  const items = search.data?.items ?? [];
  const showResults = debounced.trim().length >= GUID_SEARCH_MIN_CHARS && !tooShort;

  return (
    <div>
      <label style={{ display: "grid", gap: "0.25rem", fontSize: "0.9rem" }}>
        Найти марку
        <input
          value={text}
          onChange={(event) => setText(event.target.value)}
          style={{ border: "1px solid #d0d5dd", borderRadius: 8, padding: "0.4rem 0.55rem" }}
        />
      </label>
      {tooShort ? (
        <p style={{ margin: "0.25rem 0 0", color: "#667085" }}>
          Введите не меньше {GUID_SEARCH_MIN_CHARS} знаков.
        </p>
      ) : null}
      {showResults && items.length > 0 ? (
        <CandidateList items={items} testId="guid-explain-search-results" onReplace={onReplace} />
      ) : null}
      {showResults && items.length === 0 && !search.isLoading ? (
        <p style={{ margin: "0.25rem 0 0", color: "#667085" }}>Ничего не найдено.</p>
      ) : null}
    </div>
  );
};

const ReplaceDialog = ({
  source,
  productKind,
  mark,
  candidate,
  onClose,
  onCommitted,
}: {
  source: GuidExplainSource;
  productKind: string;
  mark: string;
  candidate: GuidCandidate;
  onClose: () => void;
  onCommitted?: (queueNotice: string | null) => void;
}) => {
  const cards = candidate.cards ?? [];
  const needsChoice = candidate.guid_count >= 2;
  const [chosenGuid, setChosenGuid] = useState<string | null>(null);
  const [preview, setPreview] = useState<GuidReplacePreview | null>(null);
  const [busy, setBusy] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const economist = source.scope === "economist";
  const request = useRef(0);

  const load = async (guid: string) => {
    if (!source.previewReplace) {
      return;
    }
    const ticket = request.current + 1;
    request.current = ticket;
    setBusy(true);
    setError(null);
    try {
      const result = await source.previewReplace({
        product_kind: productKind,
        mark,
        target_mark: candidate.mark,
        nomenclature_id: guid.trim() || null,
        replace_oneoff: economist,
      });
      if (ticket !== request.current) {
        return;
      }
      setPreview(result);
    } catch (err) {
      if (ticket !== request.current) {
        return;
      }
      setPreview(null);
      setError(getErrorMessage(err));
    } finally {
      if (ticket === request.current) {
        setBusy(false);
      }
    }
  };

  useEffect(() => {
    if (!needsChoice) {
      void load("");
    }
    // Один номер сервер подставляет сам, поле и кнопки не нужны.
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  const pickCard = (guid: string) => {
    setChosenGuid(guid);
    setPreview(null);
    void load(guid);
  };

  const write = async () => {
    if (!source.commitReplace || !preview) {
      return;
    }
    setBusy(true);
    setError(null);
    try {
      const result = await source.commitReplace({
        product_kind: productKind,
        mark,
        target_mark: candidate.mark,
        nomenclature_id: needsChoice ? chosenGuid : null,
        replace_oneoff: economist && preview.items.some((item) => item.oneoff_notice),
      });
      onCommitted?.(result.queue_notice);
      onClose();
    } catch (err) {
      setError(getErrorMessage(err));
      setBusy(false);
    }
  };

  return (
    <div
      role="dialog"
      aria-label="Замена марки"
      style={{
        marginTop: "0.5rem",
        display: "grid",
        gap: "0.45rem",
        border: "1px solid #d0d5dd",
        borderRadius: 10,
        padding: "0.7rem",
      }}
    >
      <p style={{ margin: 0 }}>Заменить на «{candidate.mark}»</p>
      {needsChoice ? (
        <div
          role="group"
          aria-label="Выберите карточку"
          style={{ display: "flex", gap: "0.5rem", flexWrap: "wrap" }}
        >
          {cards.map((card) => (
            <Button
              key={card.guid}
              type="button"
              variant={chosenGuid === card.guid ? "primary" : "secondary"}
              aria-pressed={chosenGuid === card.guid}
              disabled={busy}
              onClick={() => pickCard(card.guid)}
            >
              {cardButtonLabel(card, cards)}
            </Button>
          ))}
        </div>
      ) : null}
      {error ? <Alert tone="error">{error}</Alert> : null}
      {preview ? (
        <div data-testid="guid-replace-preview" style={{ display: "grid", gap: "0.45rem" }}>
          {preview.items.map((item) => (
            <div key={item.kp_id}>
              <div>
                {item.kp_id > 0
                  ? `КП №${item.kp_id}, количество ${item.qty}`
                  : `Количество ${item.qty}`}
              </div>
              <div>
                Цена за штуку со скидкой: {rub(item.was.unit_discounted)} →{" "}
                {rub(item.will.unit_discounted)}
              </div>
              <div>
                Сумма строки: {rub(item.was.line_sum)} → {rub(item.will.line_sum)}
              </div>
              <div>
                Итог КП: {rub(item.was.total)} → {rub(item.will.total)}
              </div>
              {item.delivery_was !== null && item.delivery_will !== null ? (
                <div>
                  Доставка: {rub(item.delivery_was)} → {rub(item.delivery_will)}
                </div>
              ) : null}
              {item.oneoff_notice ? <div>{item.oneoff_notice}</div> : null}
            </div>
          ))}
          {preview.skipped.map((item) => (
            <div key={`skip-${item.kp_id}`}>КП №{item.kp_id}: {item.reason}</div>
          ))}
        </div>
      ) : null}
      <div style={{ display: "flex", gap: "0.5rem", flexWrap: "wrap" }}>
        {!preview && !needsChoice ? (
          <Button type="button" disabled={busy} onClick={() => void load("")}>
            Показать, что изменится
          </Button>
        ) : null}
        {preview ? (
          <Button type="button" disabled={busy || preview.items.length === 0} onClick={() => void write()}>
            Записать замену
          </Button>
        ) : null}
        <Button type="button" variant="secondary" onClick={onClose} disabled={busy}>
          Отмена
        </Button>
      </div>
    </div>
  );
};

const GuidExplainBody = ({ source, productKind, mark, onCommitted }: GuidExplainPanelProps) => {
  const explain = useGuidExplainQuery(source, productKind, mark);
  const data = explain.data;
  const [candidate, setCandidate] = useState<GuidCandidate | null>(null);
  const [notice, setNotice] = useState<string | null>(null);
  const canReplace = Boolean(source.previewReplace && source.commitReplace);
  const onReplace = canReplace
    ? (item: GuidCandidate) => {
        setNotice(null);
        setCandidate(item);
      }
    : undefined;

  if (explain.isError) {
    return <Alert tone="error">Не удалось загрузить объяснение. Попробуйте позже.</Alert>;
  }
  if (!data) {
    return <p style={{ margin: 0, color: "#667085" }}>Загрузка…</p>;
  }
  return (
    <div style={{ display: "grid", gap: "0.35rem" }}>
      {notice ? <Alert tone="warning">{notice}</Alert> : null}
      {data.reason ? (
        <p style={{ margin: 0 }}>
          Причина: {data.reason}
          {data.hint ? `. ${data.hint}` : ""}
        </p>
      ) : (
        <p style={{ margin: 0 }}>Номер 1С у этой строки есть.</p>
      )}
      {data.same_cards.length > 0 ? (
        <div>
          <h4 style={headingStyle}>Та же карточка, другая запись</h4>
          <CandidateList items={data.same_cards} testId="guid-explain-same-cards" onReplace={onReplace} />
        </div>
      ) : null}
      {data.slice.length > 0 ? (
        <div>
          <h4 style={headingStyle}>Другие карточки этого вида</h4>
          <CandidateList items={data.slice} testId="guid-explain-slice" onReplace={onReplace} />
        </div>
      ) : null}
      <GuidSearchBox source={source} productKind={productKind} onReplace={onReplace} />
      {candidate && canReplace ? (
        <ReplaceDialog
          source={source}
          productKind={productKind}
          mark={mark}
          candidate={candidate}
          onClose={() => setCandidate(null)}
          onCommitted={(queueNotice) => {
            onCommitted?.();
            setNotice(queueNotice);
          }}
        />
      ) : null}
    </div>
  );
};

/** Раскрывающаяся панель «Почему нет номера»: данные грузятся только после раскрытия. */
export const GuidExplainPanel = (props: GuidExplainPanelProps) => {
  const [open, setOpen] = useState(false);
  return (
    <div style={{ marginTop: "0.25rem" }}>
      <Button type="button" aria-expanded={open} onClick={() => setOpen((value) => !value)}>
        Почему нет номера и что есть рядом
      </Button>
      {open ? (
        <div style={{ marginTop: "0.5rem" }}>
          <GuidExplainBody {...props} />
        </div>
      ) : null}
    </div>
  );
};
