import type {
  CommercialDraftDetails,
  ProductType,
  SimpleKpProductType,
} from "@/features/commercial-offer/types/commercialOffer";
import {
  buildBridgePileLinesFromOrderData,
  buildBridgePilePreviewRows,
} from "@/features/commercial-offer/lib/buildBridgePilePreviewRows";
import {
  buildFbsLinesFromOrderData,
  buildFbsPreviewRows,
} from "@/features/commercial-offer/lib/buildFbsPreviewRows";
import {
  buildMarchLinesFromOrderData,
  buildMarchPreviewRows,
} from "@/features/commercial-offer/lib/buildMarchPreviewRows";
import {
  buildPileLinesFromOrderData,
  buildPilePreviewRows,
} from "@/features/commercial-offer/lib/buildPilePreviewRows";
import { buildStepLinesFromOrderData, buildStepPreviewRows } from "@/features/commercial-offer/lib/buildStepPreviewRows";
import {
  BRIDGE_PILE_GRADE_CODES,
  formatBridgePileGradeLabel,
  isBridgePileGradeCode,
} from "@/features/commercial-offer/lib/bridgePileGrades";
import { FBS_GRADE_CODES, formatFbsGradeLabel, isFbsGradeCode } from "@/features/commercial-offer/lib/fbsGrades";
import {
  MARCH_GRADE_CODES,
  formatMarchGradeLabel,
  isMarchGradeCode,
} from "@/features/commercial-offer/lib/marchGrades";
import { PILE_GRADE_CODES, formatPileGradeLabel, isPileGradeCode } from "@/features/commercial-offer/lib/pileGrades";

/**
 * Preview/grade registry per simple-KP product. Lives outside productTypeConfig on
 * purpose: the config stays a pure data module and must not import the build* / *Grades
 * modules (cyclic-import guard, plan 2026-09-08 §5). Import direction: here → config, never back.
 */
export type GradedPreviewRow = {
  lineId?: string | null;
  sourceText?: string;
  mark: string;
  name: string;
  concrete_grade: string;
  available_grades?: string[];
  qty: number;
  unit_price: number | null;
  line_total?: number | null;
  sealed?: boolean;
};

export type ProductTypePreview = {
  gradeCodes: readonly string[];
  formatGradeLabel: (code: string) => string;
  isGradeCode: (value: string) => boolean;
  buildRows: (draft: CommercialDraftDetails) => GradedPreviewRow[];
  buildLines: (rows: GradedPreviewRow[]) => string;
};

const PRODUCT_TYPE_PREVIEW: Record<SimpleKpProductType, ProductTypePreview> = {
  piles: {
    gradeCodes: PILE_GRADE_CODES,
    formatGradeLabel: formatPileGradeLabel,
    isGradeCode: isPileGradeCode,
    buildRows: buildPilePreviewRows,
    buildLines: buildPileLinesFromOrderData,
  },
  steps: {
    gradeCodes: [],
    formatGradeLabel: (code) => code,
    isGradeCode: () => false,
    buildRows: (draft) => buildStepPreviewRows(draft).map((row) => ({ ...row, concrete_grade: "" })),
    // Step order lines carry no sealed flag, so this builder has no sealed filter; it exists
    // for registry symmetry and is only reachable behind supportsGrades (i.e. never for steps).
    buildLines: buildStepLinesFromOrderData,
  },
  marches: {
    gradeCodes: MARCH_GRADE_CODES,
    formatGradeLabel: formatMarchGradeLabel,
    isGradeCode: isMarchGradeCode,
    buildRows: buildMarchPreviewRows,
    buildLines: buildMarchLinesFromOrderData,
  },
  bridge_piles: {
    gradeCodes: BRIDGE_PILE_GRADE_CODES,
    formatGradeLabel: formatBridgePileGradeLabel,
    isGradeCode: isBridgePileGradeCode,
    buildRows: buildBridgePilePreviewRows,
    buildLines: buildBridgePileLinesFromOrderData,
  },
  fbs: {
    gradeCodes: FBS_GRADE_CODES,
    formatGradeLabel: formatFbsGradeLabel,
    isGradeCode: isFbsGradeCode,
    buildRows: buildFbsPreviewRows,
    buildLines: buildFbsLinesFromOrderData,
  },
};

/** Undefined for plates — the plate preview stays a separate component. */
export const getProductTypePreview = (type: ProductType): ProductTypePreview | undefined =>
  type === "plates" ? undefined : PRODUCT_TYPE_PREVIEW[type];

/** Always defined: the registry covers every simple-KP product type. */
export const getSimpleProductTypePreview = (type: SimpleKpProductType): ProductTypePreview =>
  PRODUCT_TYPE_PREVIEW[type];
