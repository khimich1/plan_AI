import { z } from "zod";
import { EXECUTION_TERMS_PARSE_ERROR, tryNormalizeExecutionTerms } from "@/shared/lib/executionTerms";

export const plateSubmissionSchema = z
  .object({
    text: z.string().trim().default(""),
    image: z.instanceof(File).nullable().optional(),
  })
  .refine((value) => value.text.length > 0 || value.image, {
    message: "Введите текст списка плит или загрузите изображение.",
    path: ["text"],
  });

export const clientConditionsSchema = z
  .object({
    clientName: z.string().trim().min(1, "Укажите клиента."),
    conditionsMode: z.enum(["standard", "custom"]),
    deliveryConditions: z.string().trim(),
    paymentConditions: z.string().trim(),
  })
  .superRefine((value, ctx) => {
    if (value.conditionsMode === "custom" && !value.deliveryConditions) {
      ctx.addIssue({
        code: z.ZodIssueCode.custom,
        path: ["deliveryConditions"],
        message: "Укажите условия поставки.",
      });
    }
    if (value.conditionsMode === "custom" && !value.paymentConditions) {
      ctx.addIssue({
        code: z.ZodIssueCode.custom,
        path: ["paymentConditions"],
        message: "Укажите условия оплаты.",
      });
    }
  });

export const saveOfferSchema = z
  .object({
    mode: z.enum(["database", "archive", "skip"]),
    executionTermsInput: z.string().trim(),
  })
  .superRefine((value, ctx) => {
    if (value.mode === "skip") {
      return;
    }
    const trimmed = value.executionTermsInput.trim();
    if (value.mode === "archive") {
      if (!trimmed) {
        return;
      }
      if (tryNormalizeExecutionTerms(trimmed) === null) {
        ctx.addIssue({
          code: z.ZodIssueCode.custom,
          path: ["executionTermsInput"],
          message: EXECUTION_TERMS_PARSE_ERROR,
        });
      }
      return;
    }
    if (!trimmed) {
      ctx.addIssue({
        code: z.ZodIssueCode.custom,
        path: ["executionTermsInput"],
        message: "Укажите срок изготовления.",
      });
      return;
    }
    if (tryNormalizeExecutionTerms(trimmed) === null) {
      ctx.addIssue({
        code: z.ZodIssueCode.custom,
        path: ["executionTermsInput"],
        message: EXECUTION_TERMS_PARSE_ERROR,
      });
    }
  });
