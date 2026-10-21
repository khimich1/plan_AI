export const formatContractDate = (iso: string | null | undefined): string => {
  if (!iso) {
    return "";
  }
  const [year, month, day] = iso.split("-");
  if (!year || !month || !day) {
    return iso;
  }
  return `${day}.${month}.${year}`;
};
