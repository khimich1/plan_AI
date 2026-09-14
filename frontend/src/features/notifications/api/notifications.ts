import { useMutation, useQuery, useQueryClient } from "@tanstack/react-query";
import { httpClient } from "@/shared/api/httpClient";

const BASE = "/api/v1/notifications";

export type NotificationPayload = {
  kp_id?: number;
  week_start?: string;
  reason?: string;
  [key: string]: unknown;
};

export type NotificationItem = {
  id: number;
  kind: string;
  payload: NotificationPayload;
  read_at: string | null;
  created_at: string;
};

export type NotificationList = {
  items: NotificationItem[];
  unread_count: number;
};

export const notificationsApi = {
  list: (unread = false) => {
    const query = unread ? "?unread=true" : "";
    return httpClient.get<NotificationList>(`${BASE}${query}`);
  },
  markRead: (id: number) => httpClient.post<{ id: number; read_at: string }>(`${BASE}/${id}/read`),
};

export const notificationsKeys = {
  all: ["notifications"] as const,
  list: (unread = false) => ["notifications", "list", unread] as const,
};

export const useNotificationsQuery = () =>
  useQuery<NotificationList>({
    queryKey: notificationsKeys.list(),
    queryFn: () => notificationsApi.list(false),
    staleTime: 15_000,
    refetchInterval: 30_000,
  });

export const useMarkNotificationReadMutation = () => {
  const queryClient = useQueryClient();
  return useMutation({
    mutationFn: (id: number) => notificationsApi.markRead(id),
    onSuccess: () => {
      queryClient.invalidateQueries({ queryKey: notificationsKeys.all });
    },
  });
};

/** ISO week start YYYY-MM-DD → «7.09». */
export function formatWeekStart(iso: string | undefined): string {
  if (!iso) return "—";
  const match = /^(\d{4})-(\d{2})-(\d{2})$/.exec(iso);
  if (!match) return iso;
  return `${Number(match[3])}.${match[2]}`;
}

export function formatNotificationTitle(item: NotificationItem): string {
  const kpId = item.payload.kp_id;
  if (item.kind === "promise_excluded" && typeof kpId === "number") {
    const week = formatWeekStart(item.payload.week_start);
    const reason = String(item.payload.reason ?? "").trim();
    const tail = reason ? `: ${reason}` : "";
    return `КП №${kpId} снято с недели ${week}${tail}`;
  }
  if (item.kind === "promised_date_shifted" && typeof kpId === "number") {
    const next = String(item.payload.new_promised_date ?? "").trim();
    return next
      ? `КП №${kpId}: срок обещания сдвинут на ${next}`
      : `КП №${kpId}: срок обещания сдвинут`;
  }
  return item.kind;
}

export function archiveHrefForNotification(item: NotificationItem): string | null {
  const kpId = item.payload.kp_id;
  return typeof kpId === "number" && kpId > 0 ? `/archive?kp=${kpId}` : null;
}
