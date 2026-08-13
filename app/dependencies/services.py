"""FastAPI dependency factories for application services (A6).

Override in tests via ``app.dependency_overrides``::

    from app.dependencies.services import get_archive_service

    app.dependency_overrides[get_archive_service] = lambda: fake_service

See ``tests/test_archive_endpoints.py`` for a working example.
"""
from __future__ import annotations

from fastapi import Depends

from app.dependencies.auth import get_auth_repository
from app.repositories.auth_repository import AuthRepository
from app.services.admin_service import AdminService
from app.services.auth_service import AuthService
from app.services.archive_service import ArchiveService
from app.services.commercial_calculation_service import CommercialCalculationService
from app.services.commercial_service import CommercialService
from app.services.commercial_wizard_step_service import CommercialWizardStepService
from app.services.commercial_workflow_service import CommercialWorkflowService
from app.services.delivery_schedule_service import DeliveryScheduleService
from app.services.draft_store import DraftStore
from app.services.offers_service import OffersService
from app.services.production_planning_service import ProductionPlanningService
from app.services.production_service import ProductionService
from app.services.sgp_service import SgpService
from app.services.carrier_service import CarrierService
from app.services.shipment_service import ShipmentService


def get_production_planning_service() -> ProductionPlanningService:
    return ProductionPlanningService()


def get_production_service() -> ProductionService:
    planning_service = get_production_planning_service()
    return ProductionService(
        plan_repository=planning_service.plan_repository,
        planning_service=planning_service,
    )


def get_sgp_service() -> SgpService:
    from app.repositories.kp_repository import KpRepository

    return SgpService(db_path=KpRepository().db_path)


def get_shipment_service() -> ShipmentService:
    from app.repositories.kp_repository import KpRepository
    from app.repositories.shipment_repository import ShipmentRepository

    db_path = KpRepository().db_path
    repo = ShipmentRepository(db_path=db_path)
    return ShipmentService(db_path=db_path, repo=repo)


def get_carrier_service() -> CarrierService:
    from app.repositories.kp_repository import KpRepository

    return CarrierService(db_path=KpRepository().db_path)


def get_commercial_service() -> CommercialService:
    return CommercialService()


def get_offers_service() -> OffersService:
    return OffersService()


def get_commercial_workflow_service() -> CommercialWorkflowService:
    return CommercialWorkflowService()


def get_commercial_wizard_step_service() -> CommercialWizardStepService:
    return CommercialWizardStepService(
        calculation_service=CommercialCalculationService(),
        draft_store=DraftStore(),
    )


def get_admin_service() -> AdminService:
    return AdminService()


def get_archive_service() -> ArchiveService:
    return ArchiveService()


def get_delivery_schedule_service() -> DeliveryScheduleService:
    from app.repositories.kp_repository import KpRepository

    return DeliveryScheduleService(db_path=KpRepository().db_path)


def get_auth_service(
    repository: AuthRepository = Depends(get_auth_repository),
) -> AuthService:
    return AuthService(repository=repository)
