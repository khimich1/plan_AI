import { useEffect, useId, useRef, useState, type CSSProperties } from "react";
import { useNavigate } from "react-router";
import {
  archiveHrefForNotification,
  formatNotificationTitle,
  useMarkNotificationReadMutation,
  useNotificationsQuery,
  type NotificationItem,
} from "@/features/notifications/api/notifications";

const wrapStyle: CSSProperties = {
  position: "relative",
  display: "inline-flex",
};

const buttonStyle: CSSProperties = {
  position: "relative",
  width: 36,
  height: 36,
  borderRadius: "50%",
  border: "1px solid #d6defa",
  background: "#ffffff",
  color: "#23366f",
  cursor: "pointer",
  display: "inline-flex",
  alignItems: "center",
  justifyContent: "center",
};

const badgeStyle: CSSProperties = {
  position: "absolute",
  top: -4,
  right: -4,
  minWidth: 18,
  height: 18,
  padding: "0 5px",
  borderRadius: 999,
  background: "#b42318",
  color: "#ffffff",
  fontSize: 11,
  fontWeight: 700,
  lineHeight: "18px",
  textAlign: "center",
};

const popoverStyle: CSSProperties = {
  position: "absolute",
  top: "calc(100% + 8px)",
  right: 0,
  width: 320,
  maxHeight: 360,
  overflowY: "auto",
  background: "#ffffff",
  border: "1px solid #e4e7ec",
  borderRadius: 12,
  boxShadow: "0 8px 24px rgba(15, 23, 42, 0.12)",
  zIndex: 20,
};

const emptyStyle: CSSProperties = {
  margin: 0,
  padding: "1rem",
  color: "#667085",
  fontSize: "0.875rem",
};

const listStyle: CSSProperties = {
  listStyle: "none",
  margin: 0,
  padding: 0,
};

const itemButtonStyle = (unread: boolean): CSSProperties => ({
  display: "block",
  width: "100%",
  textAlign: "left",
  padding: "0.75rem 0.9rem",
  border: "none",
  borderBottom: "1px solid #f2f4f7",
  background: unread ? "#f8fafc" : "#ffffff",
  color: "#101828",
  fontSize: "0.85rem",
  fontWeight: unread ? 600 : 400,
  cursor: "pointer",
  lineHeight: 1.4,
});

const BellIcon = () => (
  <svg
    width="18"
    height="18"
    viewBox="0 0 24 24"
    fill="none"
    stroke="currentColor"
    strokeWidth="2"
    strokeLinecap="round"
    strokeLinejoin="round"
    aria-hidden
  >
    <path d="M6 8a6 6 0 0 1 12 0c0 7 3 9 3 9H3s3-2 3-9" />
    <path d="M10.3 21a1.94 1.94 0 0 0 3.4 0" />
  </svg>
);

export const NotificationBell = () => {
  const [open, setOpen] = useState(false);
  const wrapRef = useRef<HTMLDivElement>(null);
  const navigate = useNavigate();
  const listId = useId();
  const query = useNotificationsQuery();
  const markRead = useMarkNotificationReadMutation();

  const unread = query.data?.unread_count ?? 0;
  const items = query.data?.items ?? [];

  useEffect(() => {
    if (!open) return;
    const onDocClick = (event: MouseEvent) => {
      if (wrapRef.current && !wrapRef.current.contains(event.target as Node)) {
        setOpen(false);
      }
    };
    const onKey = (event: KeyboardEvent) => {
      if (event.key === "Escape") setOpen(false);
    };
    document.addEventListener("mousedown", onDocClick);
    document.addEventListener("keydown", onKey);
    return () => {
      document.removeEventListener("mousedown", onDocClick);
      document.removeEventListener("keydown", onKey);
    };
  }, [open]);

  const onItemActivate = (item: NotificationItem) => {
    if (!item.read_at) {
      markRead.mutate(item.id);
    }
    setOpen(false);
    const href = archiveHrefForNotification(item);
    if (href) {
      navigate(href);
    }
  };

  const label = unread > 0 ? `Уведомления, непрочитанных: ${unread}` : "Уведомления";

  return (
    <div ref={wrapRef} style={wrapStyle}>
      <button
        type="button"
        aria-label={label}
        aria-expanded={open}
        aria-controls={listId}
        data-testid="notification-bell"
        onClick={() => setOpen((value) => !value)}
        style={buttonStyle}
      >
        <BellIcon />
        {unread > 0 ? (
          <span data-testid="notification-badge" style={badgeStyle}>
            {unread > 99 ? "99+" : unread}
          </span>
        ) : null}
      </button>
      {open ? (
        <div
          id={listId}
          role="dialog"
          aria-label="Уведомления"
          data-testid="notification-popover"
          style={popoverStyle}
        >
          {items.length === 0 ? (
            <p style={emptyStyle}>Нет уведомлений</p>
          ) : (
            <ul style={listStyle}>
              {items.map((item) => (
                <li key={item.id}>
                  <button
                    type="button"
                    style={itemButtonStyle(!item.read_at)}
                    onClick={() => onItemActivate(item)}
                  >
                    {formatNotificationTitle(item)}
                  </button>
                </li>
              ))}
            </ul>
          )}
        </div>
      ) : null}
    </div>
  );
};
