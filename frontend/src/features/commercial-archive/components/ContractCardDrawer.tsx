import { useEffect, useRef, useState } from "react";
import type { DragEvent, FormEvent, WheelEvent } from "react";
import { Alert } from "@/shared/ui/Alert";
import { Button } from "@/shared/ui/Button";
import { FieldWrapper, Input } from "@/shared/ui/Field";
import { Modal } from "@/shared/ui/Modal";
import { getErrorMessage } from "@/shared/lib/apiError";
import {
  useDownloadContractDocumentMutation,
  useDownloadContractEdoAgreementMutation,
  useDeleteContractCardMutation,
  useParseContractCardMutation,
  usePatchContractCardMutation,
  usePatchSupplyContractMutation,
  useSaveContractRequisitesMutation,
} from "@/features/commercial-archive/hooks/useArchiveQueries";
import { formatContractDate } from "@/features/commercial-archive/lib/formatContractDate";
import {
  blankFormFromCard,
  digitWarningText,
  formWithBank,
  mergeParsedFields,
  requisitesPayload,
} from "@/features/commercial-archive/lib/contractBlank";
import {
  SUPPLY_CONTRACT_SCAN_NOTES,
  SUPPLY_CONTRACT_STATUSES,
  type ContractCard,
  type ContractCardPatchPayload,
  type SupplyContract,
  type SupplyContractBankBundle,
  type SupplyContractFormValues,
  type SupplyContractParseResult,
} from "@/features/commercial-archive/types/supplyContract";
import { poaFieldsReady, SupplyContractFields } from "./SupplyContractFields";

export type ContractCardIntake = {
  file: File;
  parsed: SupplyContractParseResult;
};

type Props = {
  card: ContractCard;
  onUpdated?: (card: ContractCard) => void;
  onDeleted?: () => void;
  intake?: ContractCardIntake | null;
};

const DELETE_CONFIRMATION_WORD = "УДАЛИТЬ";

const fieldGrid = { display: "grid", gap: "0.45rem" } as const;
const BANK_DOUBT_KEYS = ["bank_name", "account", "corr_account", "bik"];

const IMAGE_ZOOM_MIN = 0.5;
const IMAGE_ZOOM_MAX = 3;
const IMAGE_ZOOM_STEP = 0.25;
const IMAGE_ZOOM_FIT = 1;

const clampImageZoom = (value: number) =>
  Math.min(IMAGE_ZOOM_MAX, Math.max(IMAGE_ZOOM_MIN, value));

const formatImageZoom = (zoom: number) => `${Math.round(zoom * 100)}%`;

const compactZoomButtonStyle = { minWidth: 32, padding: "0.25rem 0.5rem" } as const;

type SourceVisual = "image" | "pdf";

const isImageFile = (file: File) => file.type.startsWith("image/");

const isPdfFile = (file: File) =>
  file.type === "application/pdf" || file.name.toLowerCase().endsWith(".pdf");

const visibleSourceText = (value: string | undefined) => {
  if (!value?.trim()) {
    return null;
  }
  return value;
};

type SeededSource = {
  name: string;
  text: string | null;
  visual: SourceVisual | null;
  url: string | null;
};

const seededSource = (file: File, rawText: string | undefined): SeededSource => {
  const text = visibleSourceText(rawText);
  if (isImageFile(file)) {
    return { name: file.name, text: null, visual: "image", url: URL.createObjectURL(file) };
  }
  if (isPdfFile(file) && !text) {
    return { name: file.name, text: null, visual: "pdf", url: URL.createObjectURL(file) };
  }
  return { name: file.name, text, visual: null, url: null };
};

const previewFrameStyle = (zoom: number, visual: SourceVisual) => ({
  display: "block" as const,
  width: `${zoom * 100}%`,
  height: visual === "pdf" ? "70vh" : "auto",
  maxWidth: "none",
  border: visual === "pdf" ? 0 : undefined,
});

const ZoomControls = ({
  zoom,
  onZoom,
  onFit,
}: {
  zoom: number;
  onZoom: (direction: 1 | -1) => void;
  onFit: () => void;
}) => (
  <div style={{ display: "flex", flexWrap: "wrap", gap: "0.5rem", alignItems: "center", marginTop: "0.5rem" }}>
    <div
      style={{
        display: "inline-flex",
        flexWrap: "wrap",
        alignItems: "center",
        maxWidth: "100%",
        gap: "0.25rem",
        border: "1px solid #d0d5dd",
        borderRadius: 8,
        padding: "0.15rem",
        background: "#ffffff",
      }}
    >
      <Button
        type="button"
        variant="ghost"
        aria-label="Уменьшить"
        title="Уменьшить"
        onClick={() => onZoom(-1)}
        disabled={zoom <= IMAGE_ZOOM_MIN}
        style={compactZoomButtonStyle}
      >
        −
      </Button>
      <span style={{ minWidth: 48, textAlign: "center", fontSize: "0.85rem", color: "#475467" }}>
        {formatImageZoom(zoom)}
      </span>
      <Button
        type="button"
        variant="ghost"
        aria-label="Увеличить"
        title="Увеличить"
        onClick={() => onZoom(1)}
        disabled={zoom >= IMAGE_ZOOM_MAX}
        style={compactZoomButtonStyle}
      >
        +
      </Button>
      <Button
        type="button"
        variant="ghost"
        aria-label="По ширине"
        title="По ширине окна"
        onClick={onFit}
        style={{ padding: "0.25rem 0.5rem", fontSize: "0.85rem" }}
      >
        По ширине
      </Button>
    </div>
  </div>
);

const CardSourcePreview = ({
  name,
  text,
  url,
  visual,
  zoom,
  onZoom,
  onFit,
  onWheel,
}: {
  name: string;
  text: string | null;
  url: string | null;
  visual: SourceVisual | null;
  zoom: number;
  onZoom: (direction: 1 | -1) => void;
  onFit: () => void;
  onWheel: (event: WheelEvent<HTMLDivElement>) => void;
}) => (
  <aside aria-label="Источник карточки" style={{ position: "sticky", top: 0, alignSelf: "start", zIndex: 1 }}>
    {url && visual ? (
      <>
        <div
          onWheel={onWheel}
          style={{
            overflow: "auto",
            width: "100%",
            maxHeight: "calc(90vh - 11rem)",
            borderRadius: 12,
            border: "1px solid #e4e7ec",
            background: "#f8fafc",
          }}
        >
          {visual === "image" ? (
            <img src={url} alt="Карточка контрагента" style={previewFrameStyle(zoom, visual)} />
          ) : (
            <iframe title="Карточка контрагента" src={url} style={previewFrameStyle(zoom, visual)} />
          )}
        </div>
        <ZoomControls zoom={zoom} onZoom={onZoom} onFit={onFit} />
        <div style={{ fontSize: "0.8rem", color: "#667085" }}>Ctrl + колёсико мыши — масштаб</div>
        <div>{name}</div>
      </>
    ) : (
      <>
        <div>{name}</div>
        {text && (
          <div
            role="region"
            aria-label="Текст карточки"
            style={{
              overflow: "auto",
              whiteSpace: "pre-wrap",
              maxHeight: "calc(90vh - 11rem)",
              borderRadius: 12,
              border: "1px solid #e4e7ec",
              background: "#f8fafc",
              padding: "0.75rem",
            }}
          >
            {text}
          </div>
        )}
      </>
    )}
  </aside>
);

const withoutDoubt = (doubtful: Set<string>, keys: string[]) => {
  if (!keys.some((key) => doubtful.has(key))) {
    return doubtful;
  }
  const next = new Set(doubtful);
  for (const key of keys) {
    next.delete(key);
  }
  return next;
};

const identityPayload = (
  name: string,
  inn: string,
  number: string,
): ContractCardPatchPayload | null => {
  const trimmedName = name.trim();
  const trimmedNumber = number.trim();
  if (!trimmedName || !trimmedNumber) {
    return null;
  }
  const trimmedInn = inn.trim();
  return {
    name: trimmedName,
    inn: trimmedInn ? trimmedInn : null,
    number: trimmedNumber,
  };
};

const DeleteContractCardConfirm = ({
  open,
  pending,
  error,
  onClose,
  onConfirm,
}: {
  open: boolean;
  pending: boolean;
  error: string | null;
  onClose: () => void;
  onConfirm: () => void;
}) => {
  const [phrase, setPhrase] = useState("");

  useEffect(() => {
    if (open) {
      setPhrase("");
    }
  }, [open]);

  const confirmed = phrase === DELETE_CONFIRMATION_WORD;
  const confirm = () => {
    if (!confirmed || pending) {
      return;
    }
    onConfirm();
  };

  return (
    <Modal open={open} onClose={onClose} title="Удалить карточку?">
      <form
        aria-label="Подтверждение удаления"
        onSubmit={(event) => {
          event.preventDefault();
          confirm();
        }}
        style={{ display: "grid", gap: "0.75rem" }}
      >
        <Alert tone="warning">
          Карточка пропадёт из списка, номер снова станет свободным. Это нельзя отменить.
        </Alert>
        <p style={{ margin: 0 }}>Введите УДАЛИТЬ заглавными, без пробелов.</p>
        <FieldWrapper label="Подтверждение">
          <Input
            value={phrase}
            autoComplete="off"
            autoFocus
            onChange={(event) => setPhrase(event.target.value)}
          />
        </FieldWrapper>
        {error && <Alert tone="error">{error}</Alert>}
        <div style={{ display: "flex", justifyContent: "flex-end", gap: "0.5rem" }}>
          <Button type="button" variant="ghost" onClick={onClose} disabled={pending}>
            Отмена
          </Button>
          <Button type="button" variant="danger" disabled={!confirmed || pending} onClick={confirm}>
            {pending ? "Удаление…" : "Подтвердить удаление"}
          </Button>
        </div>
      </form>
    </Modal>
  );
};

const DownloadRow = ({
  label,
  wordLabel,
  pdfLabel,
  onWord,
  onPdf,
}: {
  label: string;
  wordLabel: string;
  pdfLabel: string;
  onWord: () => void;
  onPdf: () => void;
}) => (
  <div style={{ display: "flex", flexWrap: "wrap", gap: "0.5rem", alignItems: "center" }}>
    <span style={{ fontWeight: 600, minWidth: "11rem" }}>{label}</span>
    <Button type="button" variant="secondary" aria-label={wordLabel} onClick={onWord}>
      Word
    </Button>
    <Button type="button" variant="secondary" aria-label={pdfLabel} onClick={onPdf}>
      PDF
    </Button>
  </div>
);

export const ContractCardDrawer = ({ card, onUpdated, onDeleted, intake = null }: Props) => {
  const patchCard = usePatchContractCardMutation();
  const patchSupply = usePatchSupplyContractMutation();
  const saveBlank = useSaveContractRequisitesMutation();
  const removeCard = useDeleteContractCardMutation();
  const parseCard = useParseContractCardMutation();
  const documentDownload = useDownloadContractDocumentMutation();
  const edoDownload = useDownloadContractEdoAgreementMutation();
  const [seed] = useState(() =>
    intake ? seededSource(intake.file, intake.parsed.source_text) : null,
  );
  const [name, setName] = useState(card.name);
  const [inn, setInn] = useState(card.inn ?? "");
  const [number, setNumber] = useState(card.number);
  const [contractDate, setContractDate] = useState(card.contract_date);
  const [status, setStatus] = useState(card.status);
  const [scanNote, setScanNote] = useState(card.scan_note ?? "");
  const [hasScan, setHasScan] = useState(card.has_scan);
  const [requisitesComplete, setRequisitesComplete] = useState(card.requisites_complete);
  const [form, setForm] = useState<SupplyContractFormValues>(() =>
    intake ? mergeParsedFields(blankFormFromCard(card), intake.parsed) : blankFormFromCard(card),
  );
  const [doubtful, setDoubtful] = useState<Set<string>>(
    () => new Set(intake?.parsed.doubtful ?? []),
  );
  const [accounts, setAccounts] = useState<string[]>(() => intake?.parsed.accounts ?? []);
  const [banks, setBanks] = useState<SupplyContractBankBundle[]>(() => intake?.parsed.banks ?? []);
  const [error, setError] = useState<string | null>(null);
  const [confirmOpen, setConfirmOpen] = useState(false);
  const [deleteError, setDeleteError] = useState<string | null>(null);
  const [sourceUrl, setSourceUrl] = useState<string | null>(seed?.url ?? null);
  const [sourceName, setSourceName] = useState<string | null>(seed?.name ?? null);
  const [sourceText, setSourceText] = useState<string | null>(seed?.text ?? null);
  const [sourceVisual, setSourceVisual] = useState<SourceVisual | null>(seed?.visual ?? null);
  const [dragOver, setDragOver] = useState(false);
  const [imageZoom, setImageZoom] = useState(IMAGE_ZOOM_FIT);
  const sourceUrlRef = useRef<string | null>(seed?.url ?? null);

  useEffect(() => {
    setRequisitesComplete(card.requisites_complete);
  }, [card.requisites_complete]);

  useEffect(() => {
    return () => {
      if (sourceUrlRef.current) {
        URL.revokeObjectURL(sourceUrlRef.current);
        sourceUrlRef.current = null;
      }
    };
  }, []);

  const identity = identityPayload(name, inn, number);
  const blankReady = form.legal_form !== "" && poaFieldsReady(form);
  const blankTitle =
    form.legal_form === ""
      ? "Выберите вид"
      : poaFieldsReady(form)
        ? undefined
        : "Укажите номер и дату доверенности";
  const reviewWarning = digitWarningText(doubtful, accounts.length, banks.length);

  const setField = (key: keyof SupplyContractFormValues, value: string) => {
    setForm((current) => ({ ...current, [key]: value }) as SupplyContractFormValues);
    setDoubtful((current) => withoutDoubt(current, [key]));
  };

  const rememberSupply = (updated: SupplyContract) => {
    setStatus(updated.status);
    if ("scan_note" in updated) {
      setScanNote(updated.scan_note ?? "");
    }
    if (typeof updated.has_scan === "boolean") {
      setHasScan(updated.has_scan);
    }
    onUpdated?.({
      ...card,
      status: updated.status,
      scan_note: "scan_note" in updated ? (updated.scan_note ?? null) : card.scan_note,
      has_scan: typeof updated.has_scan === "boolean" ? updated.has_scan : card.has_scan,
    });
  };

  const runSupplyPatch = async (payload: { status?: string; scan_note?: string; file?: File }) => {
    setError(null);
    try {
      rememberSupply(await patchSupply.mutateAsync({ contractId: card.id, payload }));
    } catch (err) {
      setError(getErrorMessage(err));
    }
  };

  const onSubmitIdentity = async (event: FormEvent) => {
    event.preventDefault();
    if (!identity || patchCard.isPending) {
      return;
    }
    setError(null);
    try {
      const updated = await patchCard.mutateAsync({ cardId: card.id, payload: identity });
      setName(updated.name);
      setInn(updated.inn ?? "");
      setNumber(updated.number);
      // Номер на экране не пересчитывает дату: её меняет только ответ сервера.
      setContractDate(updated.contract_date);
      onUpdated?.(updated);
    } catch (err) {
      setError(getErrorMessage(err));
    }
  };

  const onSubmitBlank = async (event: FormEvent) => {
    event.preventDefault();
    const payload = requisitesPayload(form);
    if (!payload || !poaFieldsReady(form) || saveBlank.isPending) {
      return;
    }
    setError(null);
    try {
      const updated = await saveBlank.mutateAsync({ contractId: card.id, payload });
      if (updated.contract_date !== undefined) {
        setContractDate(updated.contract_date);
      }
      if (typeof updated.requisites_complete === "boolean") {
        setRequisitesComplete(updated.requisites_complete);
        onUpdated?.({
          ...card,
          requisites_complete: updated.requisites_complete,
          contract_date:
            updated.contract_date !== undefined ? updated.contract_date : card.contract_date,
        });
      }
    } catch (err) {
      setError(getErrorMessage(err));
    }
  };

  const rememberSourceUrl = (next: string | null) => {
    if (sourceUrlRef.current && sourceUrlRef.current !== next) {
      URL.revokeObjectURL(sourceUrlRef.current);
    }
    sourceUrlRef.current = next;
    setSourceUrl(next);
  };

  const bumpZoom = (direction: 1 | -1) => {
    setImageZoom((currentZoom) =>
      clampImageZoom(Number((currentZoom + direction * IMAGE_ZOOM_STEP).toFixed(2))),
    );
  };

  const handleImageWheel = (event: WheelEvent<HTMLDivElement>) => {
    if (!event.ctrlKey) {
      return;
    }
    event.preventDefault();
    bumpZoom(event.deltaY < 0 ? 1 : -1);
  };

  const stageSource = (file: File) => {
    setImageZoom(IMAGE_ZOOM_FIT);
    setSourceText(null);
    setSourceName(file.name);
    if (isImageFile(file)) {
      setSourceVisual("image");
      rememberSourceUrl(URL.createObjectURL(file));
      return;
    }
    setSourceVisual(null);
    rememberSourceUrl(null);
  };

  const revealParsedSource = (file: File, text: string | null) => {
    if (isImageFile(file)) {
      setSourceText(null);
      return;
    }
    // PDF без текста — скан: та же колонка масштаба, что у фото.
    if (isPdfFile(file) && !text) {
      setSourceVisual("pdf");
      setSourceText(null);
      rememberSourceUrl(URL.createObjectURL(file));
      return;
    }
    setSourceVisual(null);
    setSourceText(text);
    rememberSourceUrl(null);
  };

  const onFile = async (file: File | undefined) => {
    if (!file) {
      return;
    }
    setError(null);
    stageSource(file);
    try {
      const parsed = await parseCard.mutateAsync({ contractId: card.id, file });
      revealParsedSource(file, visibleSourceText(parsed.source_text));
      setForm((current) => mergeParsedFields(current, parsed));
      setDoubtful(new Set(parsed.doubtful ?? []));
      setAccounts(parsed.accounts ?? []);
      setBanks(parsed.banks ?? []);
    } catch (err) {
      setError(getErrorMessage(err));
    }
  };

  const onDrop = (event: DragEvent<HTMLDivElement>) => {
    event.preventDefault();
    setDragOver(false);
    void onFile(event.dataTransfer.files?.[0]);
  };

  const chooseBank = (bank: SupplyContractBankBundle) => {
    setForm((current) => formWithBank(current, bank));
    setDoubtful((current) => withoutDoubt(current, BANK_DOUBT_KEYS));
  };

  const download = async (kind: "contract" | "edo", format: "docx" | "pdf") => {
    setError(null);
    try {
      if (kind === "contract") {
        await documentDownload.mutateAsync({ contractId: card.id, format });
        return;
      }
      await edoDownload.mutateAsync({ contractId: card.id, format });
    } catch (err) {
      setError(getErrorMessage(err));
    }
  };

  const confirmDelete = async () => {
    setDeleteError(null);
    try {
      await removeCard.mutateAsync(card.id);
      setConfirmOpen(false);
      onDeleted?.();
    } catch (err) {
      setDeleteError(getErrorMessage(err));
    }
  };

  return (
    <section aria-label="Блок договора" style={{ display: "grid", gap: "1rem", marginTop: "1.25rem" }}>
      <h3 style={{ margin: 0, fontSize: "1rem" }}>
        Договор
        {requisitesComplete === false ? (
          <span style={{ marginLeft: "0.5rem", color: "#b54708", fontWeight: 600 }}>не полные</span>
        ) : null}
      </h3>
      {error && <Alert tone="error">{error}</Alert>}
      <form
        aria-label="Данные карточки"
        onSubmit={(event) => void onSubmitIdentity(event)}
        style={{ display: "grid", gap: "0.75rem" }}
      >
        <FieldWrapper label="Имя контрагента">
          <Input autoComplete="off" value={name} onChange={(event) => setName(event.target.value)} />
        </FieldWrapper>
        <FieldWrapper label="ИНН контрагента">
          <Input autoComplete="off" value={inn} onChange={(event) => setInn(event.target.value)} />
        </FieldWrapper>
        <FieldWrapper label="Номер">
          <Input autoComplete="off" value={number} onChange={(event) => setNumber(event.target.value)} />
        </FieldWrapper>
        <p style={{ margin: 0 }}>Дата: {formatContractDate(contractDate) || "—"}</p>
        <div>
          <Button type="submit" disabled={!identity || patchCard.isPending}>
            Сохранить карточку
          </Button>
        </div>
      </form>
      <div style={{ display: "flex", flexWrap: "wrap", gap: "0.75rem", alignItems: "end" }}>
        <label style={fieldGrid}>
          <span style={{ fontWeight: 600 }}>Статус</span>
          <select
            aria-label="Статус"
            value={status}
            onChange={(event) => void runSupplyPatch({ status: event.target.value })}
          >
            {SUPPLY_CONTRACT_STATUSES.map((item) => (
              <option key={item} value={item}>
                {item}
              </option>
            ))}
          </select>
        </label>
        <label style={fieldGrid}>
          <span style={{ fontWeight: 600 }}>Отметка скана</span>
          <select
            aria-label="Отметка скана"
            value={scanNote}
            onChange={(event) => void runSupplyPatch({ scan_note: event.target.value })}
          >
            {SUPPLY_CONTRACT_SCAN_NOTES.map((note) => (
              <option key={note || "empty"} value={note}>
                {note || "—"}
              </option>
            ))}
          </select>
        </label>
        <label style={fieldGrid}>
          <span style={{ fontWeight: 600 }}>Файл скана</span>
          <input
            aria-label="Файл скана"
            type="file"
            accept="application/pdf,image/jpeg,image/png"
            onChange={(event) => {
              const file = event.target.files?.[0];
              if (file) {
                void runSupplyPatch({ file });
              }
            }}
          />
        </label>
      </div>
      <p style={{ margin: 0, color: "#667085" }}>{hasScan ? "файл есть" : "файла нет"}</p>
      <div style={{ display: "grid", gap: "0.5rem" }}>
        <DownloadRow
          label="Договор"
          wordLabel="Договор Word"
          pdfLabel="Договор PDF"
          onWord={() => void download("contract", "docx")}
          onPdf={() => void download("contract", "pdf")}
        />
        <DownloadRow
          label="Соглашение об ЭДО"
          wordLabel="Соглашение об ЭДО Word"
          pdfLabel="Соглашение об ЭДО PDF"
          onWord={() => void download("edo", "docx")}
          onPdf={() => void download("edo", "pdf")}
        />
      </div>
      {reviewWarning && <Alert tone="warning">{reviewWarning}</Alert>}
      <form
        aria-label="Реквизиты договора"
        onSubmit={(event) => void onSubmitBlank(event)}
        style={{ display: "grid", gap: "0.75rem" }}
      >
        <div
          style={{
            display: "grid",
            gridTemplateColumns: sourceUrl || sourceName ? "minmax(320px, 1fr) minmax(0, 1fr)" : "1fr",
            gap: "1rem",
          }}
        >
          {sourceName && (
            <CardSourcePreview
              name={sourceName}
              text={sourceText}
              url={sourceUrl}
              visual={sourceVisual}
              zoom={imageZoom}
              onZoom={bumpZoom}
              onFit={() => setImageZoom(IMAGE_ZOOM_FIT)}
              onWheel={handleImageWheel}
            />
          )}
          <div style={{ display: "grid", gap: "0.75rem" }}>
            <div
              aria-label="Зона файла карточки"
              onDragEnter={(event) => {
                event.preventDefault();
                setDragOver(true);
              }}
              onDragOver={(event) => {
                event.preventDefault();
                setDragOver(true);
              }}
              onDragLeave={(event) => {
                event.preventDefault();
                setDragOver(false);
              }}
              onDrop={onDrop}
              style={{
                border: `2px dashed ${dragOver ? "#2b5cff" : "#d0d5dd"}`,
                borderRadius: 14,
                padding: "1rem 1.25rem",
                background: dragOver ? "#eef2ff" : "#f8faff",
                color: "#344054",
              }}
            >
              <FieldWrapper label="Карточка контрагента">
                <input
                  aria-label="Файл карточки"
                  type="file"
                  accept=".docx,.pdf,.xlsx,image/jpeg,image/png,application/pdf"
                  onChange={(event) => void onFile(event.target.files?.[0])}
                />
              </FieldWrapper>
              <div style={{ fontSize: "0.9rem", color: "#667085" }}>или перетащите файл сюда</div>
            </div>
            {banks.length > 1 && (
              <div aria-label="Банки карточки" style={{ display: "grid", gap: "0.45rem" }}>
                {banks.map((bank, index) => (
                  <Button
                    key={`${bank.bik}-${bank.account}-${index}`}
                    type="button"
                    variant="secondary"
                    onClick={() => chooseBank(bank)}
                  >
                    {[bank.bank_name, bank.account].filter(Boolean).join(" ")}
                  </Button>
                ))}
              </div>
            )}
            {accounts.length > 1 && banks.length < 2 && (
              <div aria-label="Расчётные счета" style={{ display: "grid", gap: "0.45rem" }}>
                <span style={{ fontWeight: 600 }}>Расчётные счета</span>
                {accounts.map((account) => (
                  <Button key={account} type="button" variant="secondary" onClick={() => setField("account", account)}>
                    {account}
                  </Button>
                ))}
              </div>
            )}
            <SupplyContractFields form={form} doubtful={doubtful} onChange={setField} />
            <div>
              <Button type="submit" disabled={!blankReady || saveBlank.isPending} title={blankTitle}>
                Сохранить реквизиты
              </Button>
            </div>
            <div>
              <Button type="button" variant="danger" onClick={() => setConfirmOpen(true)}>
                Удалить карточку
              </Button>
            </div>
          </div>
        </div>
      </form>
      <DeleteContractCardConfirm
        open={confirmOpen}
        pending={removeCard.isPending}
        error={deleteError}
        onClose={() => {
          if (!removeCard.isPending) {
            setConfirmOpen(false);
          }
        }}
        onConfirm={() => void confirmDelete()}
      />
    </section>
  );
};
