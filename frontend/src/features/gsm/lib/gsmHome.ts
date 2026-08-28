/** Address / home-base heuristics mirrored from core.gsm.generator. */

export const normGsmAddr = (s: string): string => {
  const t = s.toLowerCase().replaceAll("ё", "е").replaceAll("улица", "ул").replaceAll("ул.", "ул");
  return t.split(/\s+/).filter(Boolean).join(" ");
};

export const isGsmHomeBase = (addr: string): boolean => {
  const n = normGsmAddr(addr);
  if (!n.includes("кузнецкая")) {
    return false;
  }
  if (n.includes("18")) {
    return true;
  }
  return !/\d/.test(n);
};

export type GsmHomeCatalogRoute = {
  id: number;
  addr_a: string;
  addr_b: string;
  km: number;
};

export const findGsmHomeTwin = (
  chosen: GsmHomeCatalogRoute,
  catalog: GsmHomeCatalogRoute[],
): GsmHomeCatalogRoute | null => {
  const twins = catalog.filter(
    (candidate) =>
      candidate.id !== chosen.id &&
      candidate.km === chosen.km &&
      normGsmAddr(candidate.addr_a) === normGsmAddr(chosen.addr_b) &&
      normGsmAddr(candidate.addr_b) === normGsmAddr(chosen.addr_a),
  );
  if (twins.length === 0) {
    return null;
  }
  return twins.reduce((best, candidate) => (candidate.id < best.id ? candidate : best));
};
