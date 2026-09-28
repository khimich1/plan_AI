import { useEffect, useMemo, useState } from "react";
import type { DragEvent, FormEvent, ReactNode, WheelEvent } from "react";
import { Alert } from "@/shared/ui/Alert";
import { Button } from "@/shared/ui/Button";
import { FieldWrapper, Input } from "@/shared/ui/Field";
import { Modal } from "@/shared/ui/Modal";
import { getErrorMessage } from "@/shared/lib/apiError";
import {
  useCreateSupplyContractMutation,
  useDownloadSupplyContractMutation,
  useParseSupplyContractMutation,
  useReplaceSupplyContractMutation,
  useSupplyContractQuery,
  useSupplyContractRegistryQuery,
} from "@/features/commercial-archive/hooks/useArchiveQueries";
import {
  CANCELLED_CONTRACT_STATUS,
  type SupplyContractBankBundle,
  type SupplyContractCreatePayload,
  type SupplyContractRegistryRow,
} from "@/features/commercial-archive/types/supplyContract";
import { SupplyContractRegistry } from "./SupplyContractRegistry";

type Props = {
  open: boolean;
  onClose: () => void;
  kpId: number;
  counterpartyId: number | null;
  customerName?: string | null;
  customerInn?: string | null;
  customerKpp?: string | null;
};

type FormState = SupplyContractCreatePayload;

const REVIEW_LABELS = [
  ["inn", "ИНН"],
  ["kpp", "КПП"],
  ["ogrn", "ОГРН"],
  ["account", "расчётный счёт"],
  ["corr_account", "корсчёт"],
  ["bik", "БИК"],
  ["legal_address", "юридический адрес"],
  ["bank_name", "банк"],
  ["signatory_position", "должность"],
  ["signatory_name", "ФИО"],
] as const;

const BANK_PHRASE_KEYS = new Set(["bank_name", "account", "corr_account", "bik"]);

const digitWarningText = (doubtful: Set<string>, accountCount: number, bankCount: number) => {
  const labels = REVIEW_LABELS.filter(([key]) => {
    if (key === "account" && accountCount > 1) {
      return false;
    }
    if (bankCount > 1 && BANK_PHRASE_KEYS.has(key)) {
      return false;
    }
    return doubtful.has(key);
  }).map(([, label]) => label);
  if (labels.length === 0) {
    return null;
  }
  return `Проверьте по карточке: ${labels.join(", ")}.`;
};

const LIST_VALUES: Record<string, readonly string[]> = {
  legal_form: ["ooo", "ao", "ip", "kfh", "person"],
  signatory_verb: ["действующего", "действующей"],
  authority_basis: ["устав", "доверенность"],
};

const acceptedField = (key: string, value: unknown) => {
  if (value == null || value === "") {
    return false;
  }
  const allowed = LIST_VALUES[key];
  if (!allowed) {
    return true;
  }
  return allowed.includes(String(value));
};

const nameLooksDoubtful = (doubtful: Set<string>, key: "full_name" | "short_name", value: string) =>
  doubtful.has(key) && value === "";

const IMAGE_ZOOM_MIN = 0.5;
const IMAGE_ZOOM_MAX = 3;
const IMAGE_ZOOM_STEP = 0.25;
const IMAGE_ZOOM_FIT = 1;

const clampImageZoom = (value: number) =>
  Math.min(IMAGE_ZOOM_MAX, Math.max(IMAGE_ZOOM_MIN, value));

const formatImageZoom = (zoom: number) => `${Math.round(zoom * 100)}%`;

const compactZoomButtonStyle = { minWidth: 32, padding: "0.25rem 0.5rem" } as const;

const emptyForm = (
  customerName?: string | null,
  customerInn?: string | null,
  customerKpp?: string | null,
): FormState => ({
  legal_form: "ooo",
  full_name: "",
  short_name: customerName ?? "",
  signatory_position: "",
  signatory_name: "",
  signatory_verb: "действующего",
  authority_basis: "устав",
  inn: customerInn ?? "",
  kpp: customerKpp ?? "",
  ogrn: "",
  legal_address: "",
  email: "",
  bank_name: "",
  account: "",
  corr_account: "",
  bik: "",
});

export const SupplyContractArchiveButton = ({
  counterpartyId,
  onClick,
}: {
  counterpartyId?: number | null;
  onClick: () => void;
}) => {
  const missing = counterpartyId == null;
  return (
    <Button
      type="button"
      variant="secondary"
      disabled={missing}
      title={missing ? "Сначала занесите контрагента из 1С" : undefined}
      onClick={onClick}
    >
      Договор
    </Button>
  );
};

export const SupplyContractDrawer = ({
  open,
  onClose,
  kpId,
  counterpartyId,
  customerName,
  customerInn,
  customerKpp,
}: Props) => {
  const contractQuery = useSupplyContractQuery(open ? kpId : null, open && counterpartyId != null);
  const registryQuery = useSupplyContractRegistryQuery(open);
  const createMutation = useCreateSupplyContractMutation();
  const replaceMutation = useReplaceSupplyContractMutation(kpId);
  const parseMutation = useParseSupplyContractMutation();
  const downloadMutation = useDownloadSupplyContractMutation();
  const [form, setForm] = useState<FormState>(() => emptyForm(customerName, customerInn, customerKpp));
  const [doubtful, setDoubtful] = useState<Set<string>>(new Set());
  const [sourceUrl, setSourceUrl] = useState<string | null>(null);
  const [sourceName, setSourceName] = useState<string | null>(null);
  const [replacing, setReplacing] = useState<SupplyContractRegistryRow | null>(null);
  const [error, setError] = useState<string | null>(null);
  const [accounts, setAccounts] = useState<string[]>([]);
  const [banks, setBanks] = useState<SupplyContractBankBundle[]>([]);
  const [sourceText, setSourceText] = useState<string | null>(null);
  const [dragOver, setDragOver] = useState(false);
  const [imageZoom, setImageZoom] = useState(IMAGE_ZOOM_FIT);

  useEffect(() => {
    setForm(emptyForm(customerName, customerInn, customerKpp));
    setDoubtful(new Set());
    setReplacing(null);
    setError(null);
    setAccounts([]);
    setBanks([]);
    setSourceText(null);
    setDragOver(false);
    setImageZoom(IMAGE_ZOOM_FIT);
    setSourceName(null);
    setSourceUrl((current) => {
      if (current) {
        URL.revokeObjectURL(current);
      }
      return null;
    });
  }, [kpId, customerName, customerInn, customerKpp]);

  const contract = contractQuery.data ?? null;
  const cancelledForCounterparty = useMemo(
    () =>
      (registryQuery.data ?? []).some(
        (row) => row.counterparty_id === counterpartyId && row.status === CANCELLED_CONTRACT_STATUS,
      ),
    [registryQuery.data, counterpartyId],
  );
  const showCreateForm = !contract && (!cancelledForCounterparty || replacing != null);
  const pending = createMutation.isPending || replaceMutation.isPending;
  const reviewWarning = digitWarningText(doubtful, accounts.length, banks.length);

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

  const setField = (key: keyof FormState, value: string) => {
    setForm((current) => ({ ...current, [key]: value }) as FormState);
    setDoubtful((current) => {
      if (!current.has(key)) {
        return current;
      }
      const next = new Set(current);
      next.delete(key);
      return next;
    });
  };

  const applyBank = (bank: SupplyContractBankBundle) => {
    setForm((current) => ({
      ...current,
      bank_name: bank.bank_name ?? "",
      account: bank.account,
      corr_account: bank.corr_account ?? "",
      bik: bank.bik,
    }));
    setDoubtful((current) => {
      const next = new Set(current);
      for (const key of ["bank_name", "account", "corr_account", "bik"]) {
        next.delete(key);
      }
      return next;
    });
  };

  const onFile = async (file: File | undefined) => {
    if (!file) {
      return;
    }
    setError(null);
    setImageZoom(IMAGE_ZOOM_FIT);
    setSourceText(null);
    setSourceName(file.name);
    setSourceUrl((current) => {
      if (current) {
        URL.revokeObjectURL(current);
      }
      return file.type.startsWith("image/") ? URL.createObjectURL(file) : null;
    });
    try {
      const parsed = await parseMutation.mutateAsync({ kpId, file });
      const listed = parsed.accounts ?? [];
      const listedBanks = parsed.banks ?? [];
      setAccounts(listed);
      setBanks(listedBanks);
      setSourceText(parsed.source_text || null);
      setForm((current) => {
        const next = {
          ...current,
          ...Object.fromEntries(
            Object.entries(parsed.fields).filter(([key, value]) => acceptedField(key, value)),
          ),
        };
        if (listed.length > 1) {
          next.account = "";
        }
        if (listedBanks.length > 1) {
          next.bank_name = "";
          next.account = "";
          next.corr_account = "";
          next.bik = "";
        }
        return next;
      });
      setDoubtful(new Set(parsed.doubtful));
    } catch (err) {
      setError(getErrorMessage(err));
    }
  };

  const onDrop = (event: DragEvent<HTMLDivElement>) => {
    event.preventDefault();
    setDragOver(false);
    void onFile(event.dataTransfer.files?.[0]);
  };

  const onSubmit = async (event: FormEvent) => {
    event.preventDefault();
    setError(null);
    const payload: SupplyContractCreatePayload = {
      ...form,
      kpp: form.kpp || null,
      signatory_position: form.signatory_position || null,
    };
    try {
      if (replacing) {
        await replaceMutation.mutateAsync({ contractId: replacing.id, payload });
        setReplacing(null);
        return;
      }
      await createMutation.mutateAsync({ kpId, payload });
    } catch (err) {
      setError(getErrorMessage(err));
    }
  };

  return (
    <Modal open={open} onClose={onClose} title="Договор поставки" maxWidth={960}>
      {contractQuery.isPending && <div>Загрузка договора…</div>}
      {contractQuery.isError && <Alert tone="error">{getErrorMessage(contractQuery.error)}</Alert>}
      {error && <Alert tone="error">{error}</Alert>}
      {showCreateForm && reviewWarning && (
        <div style={{ marginBottom: "1rem" }}>
          <Alert tone="warning">{reviewWarning}</Alert>
        </div>
      )}
      {contract && !replacing && (
        <section aria-label="Действующий договор" style={{ display: "grid", gap: "0.5rem", marginBottom: "1rem" }}>
          <div>Номер: {contract.number}</div>
          <div>Дата: {contract.contract_date}</div>
          <div>Статус: {contract.status}</div>
          <Button
            type="button"
            variant="secondary"
            disabled={downloadMutation.isPending}
            onClick={() => downloadMutation.mutate(kpId)}
          >
            Скачать docx
          </Button>
        </section>
      )}
      {showCreateForm && (
        <form onSubmit={(event) => void onSubmit(event)} style={{ marginBottom: "1.25rem" }}>
          <div
            style={{
              display: "grid",
              gridTemplateColumns: sourceUrl || sourceName ? "minmax(320px, 1fr) minmax(0, 1fr)" : "1fr",
              gap: "1rem",
            }}
          >
            {(sourceUrl || sourceName) && (
              <aside
                aria-label="Источник карточки"
                style={{ position: "sticky", top: 0, alignSelf: "start", zIndex: 1 }}
              >
                {sourceUrl ? (
                  <>
                    <div
                      onWheel={handleImageWheel}
                      style={{
                        overflow: "auto",
                        width: "100%",
                        maxHeight: "calc(90vh - 11rem)",
                        borderRadius: 12,
                        border: "1px solid #e4e7ec",
                        background: "#f8fafc",
                      }}
                    >
                      <img
                        src={sourceUrl}
                        alt="Карточка контрагента"
                        style={{
                          display: "block",
                          width: `${imageZoom * 100}%`,
                          height: "auto",
                          maxWidth: "none",
                        }}
                      />
                    </div>
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
                          onClick={() => bumpZoom(-1)}
                          disabled={imageZoom <= IMAGE_ZOOM_MIN}
                          style={compactZoomButtonStyle}
                        >
                          −
                        </Button>
                        <span style={{ minWidth: 48, textAlign: "center", fontSize: "0.85rem", color: "#475467" }}>
                          {formatImageZoom(imageZoom)}
                        </span>
                        <Button
                          type="button"
                          variant="ghost"
                          aria-label="Увеличить"
                          title="Увеличить"
                          onClick={() => bumpZoom(1)}
                          disabled={imageZoom >= IMAGE_ZOOM_MAX}
                          style={compactZoomButtonStyle}
                        >
                          +
                        </Button>
                        <Button
                          type="button"
                          variant="ghost"
                          aria-label="По ширине"
                          title="По ширине окна"
                          onClick={() => setImageZoom(IMAGE_ZOOM_FIT)}
                          style={{ padding: "0.25rem 0.5rem", fontSize: "0.85rem" }}
                        >
                          По ширине
                        </Button>
                      </div>
                    </div>
                    <div style={{ fontSize: "0.8rem", color: "#667085" }}>Ctrl + колёсико мыши — масштаб</div>
                    {sourceName && <div>{sourceName}</div>}
                  </>
                ) : (
                  <>
                    {sourceName && <div>{sourceName}</div>}
                    {sourceText && (
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
                        {sourceText}
                      </div>
                    )}
                  </>
                )}
              </aside>
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
                      onClick={() => applyBank(bank)}
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
                    <Button
                      key={account}
                      type="button"
                      variant="secondary"
                      onClick={() => setField("account", account)}
                    >
                      {account}
                    </Button>
                  ))}
                </div>
              )}
              <label style={{ display: "grid", gap: "0.45rem" }}>
                <span style={{ fontWeight: 600 }}>Вид</span>
                <select
                  aria-label="Вид"
                  value={form.legal_form}
                  onChange={(event) => setField("legal_form", event.target.value)}
                >
                  <option value="ooo">ООО</option>
                  <option value="ao">АО</option>
                  <option value="ip">ИП</option>
                  <option value="kfh">КФХ</option>
                  <option value="person">физлицо</option>
                </select>
              </label>
              <DoubtfulField
                label="Полное наименование"
                doubtful={nameLooksDoubtful(doubtful, "full_name", form.full_name)}
              >
                <Input
                  aria-label="Полное наименование"
                  value={form.full_name}
                  onChange={(event) => setField("full_name", event.target.value)}
                />
              </DoubtfulField>
              <DoubtfulField
                label="Краткое наименование"
                doubtful={nameLooksDoubtful(doubtful, "short_name", form.short_name)}
              >
                <Input
                  aria-label="Краткое наименование"
                  value={form.short_name}
                  onChange={(event) => setField("short_name", event.target.value)}
                />
              </DoubtfulField>
              <DoubtfulField label="Должность" doubtful={doubtful.has("signatory_position")}>
                <Input
                  aria-label="Должность"
                  value={form.signatory_position ?? ""}
                  onChange={(event) => setField("signatory_position", event.target.value)}
                />
              </DoubtfulField>
              <DoubtfulField label="ФИО подписанта" doubtful={doubtful.has("signatory_name")}>
                <Input
                  aria-label="ФИО подписанта"
                  value={form.signatory_name}
                  onChange={(event) => setField("signatory_name", event.target.value)}
                />
              </DoubtfulField>
              <label style={{ display: "grid", gap: "0.45rem" }}>
                <span style={{ fontWeight: 600 }}>Род</span>
                <select
                  aria-label="Род"
                  value={form.signatory_verb}
                  onChange={(event) => setField("signatory_verb", event.target.value)}
                >
                  <option value="действующего">действующего</option>
                  <option value="действующей">действующей</option>
                </select>
              </label>
              <label style={{ display: "grid", gap: "0.45rem" }}>
                <span style={{ fontWeight: 600 }}>Основание</span>
                <select
                  aria-label="Основание"
                  value={form.authority_basis}
                  onChange={(event) => setField("authority_basis", event.target.value)}
                >
                  <option value="устав">устав</option>
                  <option value="доверенность">доверенность</option>
                </select>
              </label>
              <DoubtfulField label="ИНН" doubtful={doubtful.has("inn")}>
                <Input
                  aria-label="ИНН"
                  value={form.inn}
                  onChange={(event) => setField("inn", event.target.value)}
                />
              </DoubtfulField>
              <DoubtfulField label="КПП" doubtful={doubtful.has("kpp")}>
                <Input
                  aria-label="КПП"
                  value={form.kpp ?? ""}
                  onChange={(event) => setField("kpp", event.target.value)}
                />
              </DoubtfulField>
              <DoubtfulField label="ОГРН" doubtful={doubtful.has("ogrn")}>
                <Input
                  aria-label="ОГРН"
                  value={form.ogrn}
                  onChange={(event) => setField("ogrn", event.target.value)}
                />
              </DoubtfulField>
              <DoubtfulField label="Юридический адрес" doubtful={doubtful.has("legal_address")}>
                <Input
                  aria-label="Юридический адрес"
                  value={form.legal_address}
                  onChange={(event) => setField("legal_address", event.target.value)}
                />
              </DoubtfulField>
              <DoubtfulField label="E-mail" doubtful={doubtful.has("email")}>
                <Input
                  aria-label="E-mail"
                  value={form.email}
                  onChange={(event) => setField("email", event.target.value)}
                />
              </DoubtfulField>
              <DoubtfulField label="Банк" doubtful={doubtful.has("bank_name")}>
                <Input
                  aria-label="Банк"
                  value={form.bank_name}
                  onChange={(event) => setField("bank_name", event.target.value)}
                />
              </DoubtfulField>
              <DoubtfulField label="Расчётный счёт" doubtful={doubtful.has("account")}>
                <Input
                  aria-label="Расчётный счёт"
                  value={form.account}
                  onChange={(event) => setField("account", event.target.value)}
                />
              </DoubtfulField>
              <DoubtfulField label="Корсчёт" doubtful={doubtful.has("corr_account")}>
                <Input
                  aria-label="Корсчёт"
                  value={form.corr_account}
                  onChange={(event) => setField("corr_account", event.target.value)}
                />
              </DoubtfulField>
              <DoubtfulField label="БИК" doubtful={doubtful.has("bik")}>
                <Input
                  aria-label="БИК"
                  value={form.bik}
                  onChange={(event) => setField("bik", event.target.value)}
                />
              </DoubtfulField>
              <Button type="submit" disabled={pending}>
                {pending ? "Сохраняем…" : "Подтвердить"}
              </Button>
            </div>
          </div>
        </form>
      )}
      <SupplyContractRegistry
        open={open}
        onRequestNew={(row) => {
          setReplacing(row);
          setError(null);
        }}
      />
    </Modal>
  );
};

const DoubtfulField = ({
  label,
  doubtful,
  children,
}: {
  label: string;
  doubtful: boolean;
  children: ReactNode;
}) => (
  <div
    data-doubtful={doubtful ? "true" : undefined}
    style={doubtful ? { outline: "2px solid #f5b942", background: "#fff8e8", borderRadius: 12 } : undefined}
  >
    <FieldWrapper label={label}>{children}</FieldWrapper>
  </div>
);
