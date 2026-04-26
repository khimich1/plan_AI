import type { PropsWithChildren } from "react";
import { QueryClientProvider } from "@tanstack/react-query";
import { queryClient } from "@/shared/lib/queryClient";
import { AuthProvider } from "@/features/auth/model/AuthProvider";
import { WizardDraftProvider } from "@/features/commercial-offer/store/wizardDraftStore";

export const AppProviders = ({ children }: PropsWithChildren) => (
  <QueryClientProvider client={queryClient}>
    <AuthProvider>
      <WizardDraftProvider>{children}</WizardDraftProvider>
    </AuthProvider>
  </QueryClientProvider>
);
