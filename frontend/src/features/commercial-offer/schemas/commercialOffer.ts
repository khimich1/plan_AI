import { z } from "zod";

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
    if (value.mode === "database" && !value.executionTermsInput) {
      ctx.addIssue({
        code: z.ZodIssueCode.custom,
        path: ["executionTermsInput"],
        message: "Укажите срок изготовления.",
      });
    }
  });
