import { useState } from "react";
import { NavLink, useHref, useLocation, useNavigate } from "react-router";
import { useCommercialDraftHeaderBridge } from "@/pages/commercial-offer-create/CommercialOfferHeaderBridge";
import { useAuth } from "@/features/auth/model/AuthProvider";
import { Modal } from "@/shared/ui/Modal";
import { Button } from "@/shared/ui/Button";
import { DbManagementModal } from "@/features/admin/components/DbManagementModal";

const NEW_OFFER_PATH = "/new";
const ARCHIVE_PATH = "/archive";
const PRODUCTION_PATH = "/production";
const LOGISTICS_PATH = "/logistics";
const GSM_PATH = "/gsm";

export const AppHeader = () => {
  const { hasDraft, resetDraft } = useCommercialDraftHeaderBridge();
  const { user, logout, isLoggingOut } = useAuth();
  const navigate = useNavigate();
  const location = useLocation();
  const newOfferHref = useHref(NEW_OFFER_PATH);
  const [confirmOpen, setConfirmOpen] = useState(false);
  const [dbModalOpen, setDbModalOpen] = useState(false);
  const isAdmin = user?.role === "admin";
  const isProduction = user?.role === "production";
  const isManager = user?.role === "manager";
  const isLogistics = user?.role === "logistics";
  const isAccountant = user?.role === "accountant";
  const canSeeCommercial = isAdmin || isManager;
  const canSeeProduction = isAdmin || isProduction;
  const canSeeLogistics = isAdmin || isLogistics;
  const canSeeGsm = isAdmin || isAccountant;

  const onLogoutClick = async () => {
    try {
      await logout();
    } finally {
      resetDraft();
      navigate("/login", { replace: true });
    }
  };

  const onNewOfferClick = (event: React.MouseEvent<HTMLAnchorElement>) => {
    event.preventDefault();
    if (location.pathname === NEW_OFFER_PATH) {
      return;
    }
    if (hasDraft) {
      setConfirmOpen(true);
      return;
    }
    navigate(NEW_OFFER_PATH);
  };

  const continueDraft = () => {
    setConfirmOpen(false);
    navigate(NEW_OFFER_PATH);
  };

  const onResetDraftAndGoNew = () => {
    resetDraft();
    setConfirmOpen(false);
    navigate(NEW_OFFER_PATH, { replace: true });
  };

  return (
    <header className="app-header">
      <div className="app-header__inner">
        <div className="app-header__brand">
          <span className="app-header__logo" aria-hidden>
            S
          </span>
          <div>
            <div className="app-header__title">Коммерческие предложения</div>
            <div className="app-header__subtitle">Внутренний кабинет</div>
          </div>
        </div>
        <nav className="app-nav">
          {canSeeCommercial && (
            <>
              <a
                href={newOfferHref}
                onClick={onNewOfferClick}
                className={
                  location.pathname === NEW_OFFER_PATH ? "app-nav__link app-nav__link--active" : "app-nav__link"
                }
              >
                Создать КП
              </a>
              <NavLink
                to={ARCHIVE_PATH}
                className={({ isActive }) =>
                  isActive ? "app-nav__link app-nav__link--active" : "app-nav__link"
                }
              >
                Архив
              </NavLink>
            </>
          )}
          {canSeeProduction && (
            <NavLink
              to={PRODUCTION_PATH}
              className={({ isActive }) =>
                isActive ? "app-nav__link app-nav__link--active" : "app-nav__link"
              }
            >
              Производство
            </NavLink>
          )}
          {canSeeLogistics && (
            <NavLink
              to={LOGISTICS_PATH}
              className={({ isActive }) =>
                isActive ? "app-nav__link app-nav__link--active" : "app-nav__link"
              }
            >
              Логистика
            </NavLink>
          )}
          {canSeeGsm && (
            <NavLink
              to={GSM_PATH}
              className={({ isActive }) =>
                isActive ? "app-nav__link app-nav__link--active" : "app-nav__link"
              }
            >
              ГСМ
            </NavLink>
          )}
          {isAdmin && (
            <button
              type="button"
              onClick={() => setDbModalOpen(true)}
              aria-label="Управление БД"
              title="Управление БД"
              style={{
                marginLeft: "0.75rem",
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
              }}
            >
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
                <circle cx="12" cy="12" r="3" />
                <path d="M19.4 15a1.65 1.65 0 0 0 .33 1.82l.06.06a2 2 0 1 1-2.83 2.83l-.06-.06a1.65 1.65 0 0 0-1.82-.33 1.65 1.65 0 0 0-1 1.51V21a2 2 0 0 1-4 0v-.09a1.65 1.65 0 0 0-1-1.51 1.65 1.65 0 0 0-1.82.33l-.06.06a2 2 0 1 1-2.83-2.83l.06-.06a1.65 1.65 0 0 0 .33-1.82 1.65 1.65 0 0 0-1.51-1H3a2 2 0 0 1 0-4h.09a1.65 1.65 0 0 0 1.51-1 1.65 1.65 0 0 0-.33-1.82l-.06-.06a2 2 0 1 1 2.83-2.83l.06.06a1.65 1.65 0 0 0 1.82.33H9a1.65 1.65 0 0 0 1-1.51V3a2 2 0 0 1 4 0v.09a1.65 1.65 0 0 0 1 1.51 1.65 1.65 0 0 0 1.82-.33l.06-.06a2 2 0 1 1 2.83 2.83l-.06.06a1.65 1.65 0 0 0-.33 1.82V9a1.65 1.65 0 0 0 1.51 1H21a2 2 0 0 1 0 4h-.09a1.65 1.65 0 0 0-1.51 1z" />
              </svg>
            </button>
          )}
          {user && (
            <div className="app-header__user">
              <div className="app-header__user-info">
                <span className="app-header__user-name">{user.username}</span>
                <span className="app-header__user-role">{user.role}</span>
              </div>
              <Button variant="ghost" onClick={onLogoutClick} disabled={isLoggingOut}>
                {isLoggingOut ? "Выход..." : "Выйти"}
              </Button>
            </div>
          )}
        </nav>
      </div>

      <Modal open={confirmOpen} onClose={() => setConfirmOpen(false)} title="У вас есть незавершённый черновик">
        <p style={{ marginTop: 0 }}>
          Продолжить работу с текущим черновиком КП или начать заново? При «Начать заново» все данные черновика
          будут удалены.
        </p>
        <div style={{ display: "flex", justifyContent: "flex-end", gap: "0.5rem", marginTop: "1rem" }}>
          <Button variant="ghost" onClick={() => setConfirmOpen(false)}>
            Отмена
          </Button>
          <Button variant="secondary" onClick={continueDraft}>
            Продолжить
          </Button>
          <Button variant="danger" onClick={onResetDraftAndGoNew}>
            Начать заново
          </Button>
        </div>
      </Modal>

      {isAdmin && (
        <DbManagementModal
          open={dbModalOpen}
          onClose={() => setDbModalOpen(false)}
        />
      )}
    </header>
  );
};
