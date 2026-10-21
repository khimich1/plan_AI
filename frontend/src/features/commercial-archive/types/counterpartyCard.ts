export type CounterpartyCard = {
  id: number;
  code_1c: string;
  name: string;
  inn: string | null;
  kpp: string | null;
  contract_id: number | null;
  contract_number: string | null;
  contract_status: string | null;
  contract_date: string | null;
};

export type CounterpartyCardList = {
  items: CounterpartyCard[];
  count: number;
  truncated: boolean;
};
