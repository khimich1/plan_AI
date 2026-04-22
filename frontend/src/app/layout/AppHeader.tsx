import { useState } from "react";
import { NavLink, useLocation, useNavigate } from "react-router-dom";
import { useWizardDraftStore } from "@/features/commercial-offer/store/wizardDraftStore";
import { draftStorage } from "@/features/commercial-offer/store/draftStorage";
import { Modal } from "@/shared/ui/Modal";
import { Button } from "@/shared/ui/Button";

const NEW_OFFER_PATH = "/commercial-offer/new";
const ARCHIVE_PATH = "/commercial-offer/archive";
const PRODUCTION_PATH = "/production";

const hasDraft = (state: ReturnType<typeof useWizardDraftStore>["state"]): boolean => {
  if (state.draftId) {
    return true;
  }
  if (state.sourceText && state.sourceText.trim().length > 0) {
    return true;
  }
  if (state.selectedImageName) {
    return true;
  }
  if (state.managerId || state.clientName || state.lastSaveResult) {
    return true;
  }
  return false;
};

export const AppHeader = () => {
  const { state, dispatch } = useWizardDraftStore();
  const navigate = useNavigate();
  const location = useLocation();
  const [confirmOpen, setConfirmOpen] = useState(false);

  const onNewOfferClick = (event: React.MouseEvent<HTMLAnchorElement>) => {
    event.preventDefault();
    if (location.pathname === NEW_OFFER_PATH) {
      return;
    }
    if (hasDraft(state)) {
      setConfirmOpen(true);
      return;
    }
    navigate(NEW_OFFER_PATH);
  };

  const continueDraft = () => {
    setConfirmOpen(false);
    navigate(NEW_OFFER_PATH);
  };

  const resetDraft = () => {
    dispatch({ type: "reset" });
    draftStorage.clear();
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
            <div className="app-header__subtitle">Шишов · внутренний кабинет</div>
          </div>
        </div>
        <nav className="app-nav">
          <a
            href={NEW_OFFER_PATH}
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
          <NavLink
            to={PRODUCTION_PATH}
            className={({ isActive }) =>
              isActive ? "app-nav__link app-nav__link--active" : "app-nav__link"
            }
          >
            Производство
          </NavLink>
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
          <Button variant="danger" onClick={resetDraft}>
            Начать заново
          </Button>
        </div>
      </Modal>
    </header>
  );
};
