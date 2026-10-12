import { useEffect, useRef, useState } from "react";
import { Alert } from "@/shared/ui/Alert";
import { Button } from "@/shared/ui/Button";
import { PileCostKnobs } from "@/features/price-desk/components/PileCostKnobs";
import { PileCostMarkCard } from "@/features/price-desk/components/PileCostMarkCard";
import {
  buildInputs,
  draftFromInputs,
  formatPileAmount,
  type KnobDraft,
} from "@/features/price-desk/components/pileCostFormat";
import {
  usePileCostApplyMutation,
  usePileCostPreviewMutation,
  usePileCostQuery,
  useSavePileCostInputsMutation,
} from "@/features/price-desk/hooks/usePileCostQueries";
import type { PileCostPreview, PileCostState } from "@/features/price-desk/types/pileCost";
import { getErrorMessage } from "@/shared/lib/apiError";

const CONSTANTS =
  "Коэффициенты только текстом: надбавка к расходам 3%, наценка ×1,2, НДС ×1,22, для сечений 30, 35 и 40 см ещё ×1,02.";

export const PileCostPanel = () => {
  const [selectedMark, setSelectedMark] = useState<string | null>(null);
  const [markQuery, setMarkQuery] = useState("");
  const query = usePileCostQuery(selectedMark);
  const saveMutation = useSavePileCostInputsMutation();
  const previewMutation = usePileCostPreviewMutation();
  const applyMutation = usePileCostApplyMutation();

  const [snapshot, setSnapshot] = useState<PileCostState | null>(null);
  const [draft, setDraft] = useState<KnobDraft | null>(null);
  const [adoptedSha, setAdoptedSha] = useState<string | null>(null);
  const [dirty, setDirty] = useState(false);
  const [preview, setPreview] = useState<PileCostPreview | null>(null);
  const [synced, setSynced] = useState(false);
  const [notice, setNotice] = useState<string | null>(null);
  const [localError, setLocalError] = useState<string | null>(null);
  const lastSent = useRef("");

  useEffect(() => {
    if (query.data) {
      setSnapshot(query.data);
    }
  }, [query.data]);

  const data = query.data ?? snapshot;
  const effectiveDraft =
    draft && data && (dirty || adoptedSha === data.inputs_sha256)
      ? draft
      : data
        ? draftFromInputs(data.inputs)
        : null;

  useEffect(() => {
    if (!data || dirty || data.inputs_sha256 === adoptedSha) {
      return;
    }
    setDraft(draftFromInputs(data.inputs));
    setAdoptedSha(data.inputs_sha256);
  }, [data, dirty, adoptedSha]);

  const busy = saveMutation.isPending || previewMutation.isPending || applyMutation.isPending;
  const errorText = localError ?? (query.isError ? getErrorMessage(query.error) : null);
  const breakdown =
    query.data?.breakdown && query.data.breakdown.mark === selectedMark ? query.data.breakdown : null;

  const forgetDiff = () => {
    setPreview(null);
    setSynced(false);
    setNotice(null);
    previewMutation.reset();
    applyMutation.reset();
  };

  const editDraft = (next: KnobDraft) => {
    setDraft(next);
    setDirty(true);
    setLocalError(null);
    forgetDiff();
  };

  const persist = async (next: KnobDraft) => {
    if (!data) {
      return false;
    }
    setDraft(next);
    const built = buildInputs(next, data.inputs);
    if (!built.ok) {
      setDirty(true);
      setLocalError(built.error);
      return false;
    }
    const payload = JSON.stringify(built.body);
    if (payload === JSON.stringify(data.inputs) || payload === lastSent.current) {
      setDirty(false);
      setLocalError(null);
      return true;
    }
    setLocalError(null);
    forgetDiff();
    try {
      const saved = await saveMutation.mutateAsync(built.body);
      lastSent.current = JSON.stringify(saved.inputs);
      setDirty(false);
      return true;
    } catch (err) {
      setDirty(true);
      setLocalError(getErrorMessage(err));
      return false;
    }
  };

  const runPreview = async () => {
    if (dirty && effectiveDraft) {
      const saved = await persist(effectiveDraft);
      if (!saved) {
        return;
      }
    }
    setLocalError(null);
    try {
      const result = await previewMutation.mutateAsync();
      setPreview(result);
      const already = result.changed === 0 && result.new === 0;
      setSynced(already);
      setNotice(already ? "Цены уже совпадают с прайсом." : null);
    } catch (err) {
      setPreview(null);
      setLocalError(getErrorMessage(err));
    }
  };

  const runApply = async () => {
    if (!preview) {
      return;
    }
    setLocalError(null);
    try {
      await applyMutation.mutateAsync(preview.inputs_sha256);
      setPreview(null);
      setSynced(true);
      setNotice("Записано в прайс.");
    } catch (err) {
      setLocalError(getErrorMessage(err));
    }
  };

  return (
    <section aria-label="Цельные сваи" style={{ display: "grid", gap: "0.85rem" }}>
      <header>
        <h2 style={{ margin: 0, fontSize: "1.1rem" }}>Цельные сваи</h2>
        <p style={{ margin: "0.35rem 0 0", color: "#475467" }}>{CONSTANTS}</p>
      </header>

      {!data || !effectiveDraft ? (
        <p style={{ margin: 0, color: "#475467" }}>
          {query.isLoading ? "Загружаю ручки…" : "Ручки не загружены."}
        </p>
      ) : (
        <>
          {!synced && <Alert tone="warning">не записано в прайс</Alert>}
          {notice && <Alert tone="success">{notice}</Alert>}
          {errorText && <Alert tone="error">{errorText}</Alert>}

          <PileCostKnobs
            source={data.inputs}
            draft={effectiveDraft}
            onChange={editDraft}
            onCommit={(next) => {
              void persist(next);
            }}
          />

          <PileCostMarkCard
            marks={data.marks}
            query={markQuery}
            selectedMark={selectedMark}
            breakdown={breakdown}
            loading={query.isFetching}
            onQuery={(value) => {
              setMarkQuery(value);
              if (value.trim() !== selectedMark) {
                setSelectedMark(null);
              }
            }}
            onSelect={(mark) => {
              setSelectedMark(mark);
              setMarkQuery(mark);
            }}
          />

          {preview && <PileCostDiffView preview={preview} />}

          <div style={{ display: "flex", flexWrap: "wrap", justifyContent: "flex-end", gap: "0.5rem" }}>
            <Button type="button" variant="secondary" onClick={() => void runPreview()} disabled={busy}>
              {previewMutation.isPending ? "Считаю…" : "Показать, что изменится"}
            </Button>
            <Button type="button" variant="ghost" onClick={forgetDiff} disabled={!preview || busy}>
              Отмена
            </Button>
            <Button type="button" onClick={() => void runApply()} disabled={!preview || busy}>
              {applyMutation.isPending ? "Записываю…" : "Записать в прайс"}
            </Button>
          </div>
        </>
      )}
      {errorText && (!data || !effectiveDraft) && <Alert tone="error">{errorText}</Alert>}
    </section>
  );
};

const PileCostDiffView = ({ preview }: { preview: PileCostPreview }) => (
  <div aria-label="Дифф цельных свай" style={{ display: "grid", gap: "0.65rem" }}>
    <Alert tone="info">
      {`Изменено ${preview.changed}, новых ${preview.new}, нет в норме ${preview.missing}, без изменений ${preview.unchanged}. Марки вне нормы остаются со старой ценой.`}
    </Alert>
    <DiffList title="Изменились" items={preview.examples.changed} />
    <DiffList title="Новые" items={preview.examples.new} />
    <DiffList title="Нет в норме, цена не меняется" items={preview.examples.missing} />
  </div>
);

const DiffList = ({
  title,
  items,
}: {
  title: string;
  items: PileCostPreview["examples"]["changed"];
}) => {
  if (items.length === 0) {
    return null;
  }
  return (
    <div>
      <h3 style={{ margin: "0 0 0.35rem", fontSize: "0.95rem" }}>{title}</h3>
      <ul style={{ margin: 0, paddingLeft: "1.1rem", color: "#344054" }}>
        {items.map((item) => (
          <li key={`${title}-${item.key}`}>
            {item.key}: {formatPileAmount(item.old_price)} → {formatPileAmount(item.new_price)}
          </li>
        ))}
      </ul>
    </div>
  );
};
