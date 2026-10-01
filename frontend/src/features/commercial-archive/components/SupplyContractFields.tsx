import type { ReactNode } from "react";
import { FieldWrapper, Input } from "@/shared/ui/Field";
import type { SupplyContractCreatePayload } from "@/features/commercial-archive/types/supplyContract";

type FieldKey = keyof SupplyContractCreatePayload;

type Props = {
  form: SupplyContractCreatePayload;
  doubtful: Set<string>;
  onChange: (key: FieldKey, value: string) => void;
};

export const poaFieldsReady = (form: SupplyContractCreatePayload) => {
  if (form.authority_basis !== "доверенность") {
    return true;
  }
  return Boolean(form.poa_number?.trim() && form.poa_date?.trim());
};

const nameLooksDoubtful = (doubtful: Set<string>, key: "full_name" | "short_name", value: string) =>
  doubtful.has(key) && value === "";

export const SupplyContractFields = ({ form, doubtful, onChange }: Props) => (
  <>
    <label style={{ display: "grid", gap: "0.45rem" }}>
      <span style={{ fontWeight: 600 }}>Вид</span>
      <select aria-label="Вид" value={form.legal_form} onChange={(event) => onChange("legal_form", event.target.value)}>
        <option value="ooo">ООО</option>
        <option value="ao">АО</option>
        <option value="ip">ИП</option>
        <option value="kfh">КФХ</option>
        <option value="person">физлицо</option>
      </select>
    </label>
    <DoubtfulField label="Полное наименование" doubtful={nameLooksDoubtful(doubtful, "full_name", form.full_name)}>
      <Input
        aria-label="Полное наименование"
        value={form.full_name}
        onChange={(event) => onChange("full_name", event.target.value)}
      />
    </DoubtfulField>
    <DoubtfulField label="Краткое наименование" doubtful={nameLooksDoubtful(doubtful, "short_name", form.short_name)}>
      <Input
        aria-label="Краткое наименование"
        value={form.short_name}
        onChange={(event) => onChange("short_name", event.target.value)}
      />
    </DoubtfulField>
    <DoubtfulField label="Должность" doubtful={doubtful.has("signatory_position")}>
      <Input
        aria-label="Должность"
        value={form.signatory_position ?? ""}
        onChange={(event) => onChange("signatory_position", event.target.value)}
      />
    </DoubtfulField>
    <DoubtfulField label="ФИО подписанта" doubtful={doubtful.has("signatory_name")}>
      <Input
        aria-label="ФИО подписанта"
        value={form.signatory_name}
        onChange={(event) => onChange("signatory_name", event.target.value)}
      />
    </DoubtfulField>
    <label style={{ display: "grid", gap: "0.45rem" }}>
      <span style={{ fontWeight: 600 }}>Род</span>
      <select
        aria-label="Род"
        value={form.signatory_verb}
        onChange={(event) => onChange("signatory_verb", event.target.value)}
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
        onChange={(event) => onChange("authority_basis", event.target.value)}
      >
        <option value="устав">устав</option>
        <option value="доверенность">доверенность</option>
      </select>
    </label>
    {form.authority_basis === "доверенность" && (
      <>
        <DoubtfulField label="Номер доверенности" doubtful={doubtful.has("poa_number")}>
          <Input
            aria-label="Номер доверенности"
            value={form.poa_number ?? ""}
            onChange={(event) => onChange("poa_number", event.target.value)}
          />
        </DoubtfulField>
        <DoubtfulField label="Дата доверенности" doubtful={doubtful.has("poa_date")}>
          <Input
            aria-label="Дата доверенности"
            type="text"
            placeholder="01.03.2026"
            value={form.poa_date ?? ""}
            onChange={(event) => onChange("poa_date", event.target.value)}
          />
        </DoubtfulField>
      </>
    )}
    <DoubtfulField label="ИНН" doubtful={doubtful.has("inn")}>
      <Input aria-label="ИНН" value={form.inn} onChange={(event) => onChange("inn", event.target.value)} />
    </DoubtfulField>
    <DoubtfulField label="ОКВЭД" doubtful={doubtful.has("okved")}>
      <Input
        aria-label="ОКВЭД"
        value={form.okved ?? ""}
        onChange={(event) => onChange("okved", event.target.value)}
      />
    </DoubtfulField>
    <DoubtfulField label="КПП" doubtful={doubtful.has("kpp")}>
      <Input aria-label="КПП" value={form.kpp ?? ""} onChange={(event) => onChange("kpp", event.target.value)} />
    </DoubtfulField>
    <DoubtfulField label="ОГРН" doubtful={doubtful.has("ogrn")}>
      <Input aria-label="ОГРН" value={form.ogrn} onChange={(event) => onChange("ogrn", event.target.value)} />
    </DoubtfulField>
    <DoubtfulField label="Юридический адрес" doubtful={doubtful.has("legal_address")}>
      <Input
        aria-label="Юридический адрес"
        value={form.legal_address}
        onChange={(event) => onChange("legal_address", event.target.value)}
      />
    </DoubtfulField>
    <DoubtfulField label="E-mail" doubtful={doubtful.has("email")}>
      <Input aria-label="E-mail" value={form.email} onChange={(event) => onChange("email", event.target.value)} />
    </DoubtfulField>
    <DoubtfulField label="Банк" doubtful={doubtful.has("bank_name")}>
      <Input
        aria-label="Банк"
        value={form.bank_name}
        onChange={(event) => onChange("bank_name", event.target.value)}
      />
    </DoubtfulField>
    <DoubtfulField label="Расчётный счёт" doubtful={doubtful.has("account")}>
      <Input
        aria-label="Расчётный счёт"
        value={form.account}
        onChange={(event) => onChange("account", event.target.value)}
      />
    </DoubtfulField>
    <DoubtfulField label="Корсчёт" doubtful={doubtful.has("corr_account")}>
      <Input
        aria-label="Корсчёт"
        value={form.corr_account}
        onChange={(event) => onChange("corr_account", event.target.value)}
      />
    </DoubtfulField>
    <DoubtfulField label="БИК" doubtful={doubtful.has("bik")}>
      <Input aria-label="БИК" value={form.bik} onChange={(event) => onChange("bik", event.target.value)} />
    </DoubtfulField>
  </>
);

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
