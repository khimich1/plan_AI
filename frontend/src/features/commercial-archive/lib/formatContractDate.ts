export const formatContractDate = (iso: string): string => {
  const [year, month, day] = iso.split("-");
  if (!year || !month || !day) {
    return iso;
  }
  return `${day}.${month}.${year}`;
};
