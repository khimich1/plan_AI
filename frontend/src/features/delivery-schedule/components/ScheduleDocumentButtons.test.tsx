import { cleanup, fireEvent, render, screen, waitFor } from "@testing-library/react";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { ScheduleDocumentButtons } from "@/features/delivery-schedule/components/ScheduleDocumentButtons";
import type { BatchDraft, OfferPlateForSchedule } from "@/features/delivery-schedule/lib/scheduleDraft";

const templateMutate = vi.fn();
const documentMutate = vi.fn();
const saveBlobAs = vi.fn();
const buildScheduleTemplateXlsx = vi.fn();

vi.mock("@/features/delivery-schedule/hooks/useDeliveryScheduleQueries", () => ({
  useDownloadDeliveryScheduleTemplateMutation: () => ({
    mutateAsync: templateMutate,
    isPending: false,
  }),
  useDownloadDeliveryScheduleDocumentMutation: () => ({
    mutateAsync: documentMutate,
    isPending: false,
    variables: undefined,
  }),
}));

vi.mock("@/shared/lib/downloadFile", () => ({
  saveBlobAs: (...args: unknown[]) => saveBlobAs(...args),
}));

vi.mock("@/features/delivery-schedule/lib/buildScheduleTemplateXlsx", () => ({
  TEMPLATE_FILENAME: "delivery_schedule_template.xlsx",
  buildScheduleTemplateXlsx: (...args: unknown[]) => buildScheduleTemplateXlsx(...args),
}));

const plates: OfferPlateForSchedule[] = [{ id: 1, plate_name: "ПБ 60-12-8", qty: 40 }];
const batches: BatchDraft[] = [
  {
    key: "k1",
    name: "Партия 1",
    deliver_from: "2026-04-01",
    deliver_to: "2026-04-10",
    produce_by: "2026-03-25",
    items: [{ plate_id: 1, qty: 10 }],
    status: null,
    ready_date: null,
    hint: null,
    changed: false,
  },
];

describe("ScheduleDocumentButtons", () => {
  afterEach(() => {
    cleanup();
  });

  beforeEach(() => {
    templateMutate.mockReset();
    documentMutate.mockReset();
    saveBlobAs.mockReset();
    buildScheduleTemplateXlsx.mockReset();
    buildScheduleTemplateXlsx.mockResolvedValue(new Blob(["xlsx"]));
  });

  it("builds the template on the client from the open editor draft", async () => {
    render(
      <ScheduleDocumentButtons kpId={7} plates={plates} batches={batches} />,
    );

    fireEvent.click(screen.getByRole("button", { name: "Скачать шаблон" }));

    await waitFor(() => {
      expect(buildScheduleTemplateXlsx).toHaveBeenCalledWith(plates, batches);
    });
    expect(saveBlobAs).toHaveBeenCalledWith(expect.any(Blob), "delivery_schedule_template.xlsx");
    expect(templateMutate).not.toHaveBeenCalled();
  });
});
