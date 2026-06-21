"""FastAPI dependency factories for application services (A6).

Override in tests via ``app.dependency_overrides``::

    from app.dependencies.services import get_archive_service

    app.dependency_overrides[get_archive_service] = lambda: fake_service

See ``tests/test_archive_endpoints.py`` for a working example.
"""
from __future__ import annotations

from app.services.admin_service import AdminService
from app.services.archive_service import ArchiveService
from app.services.commercial_service import CommercialService
from app.services.commercial_workflow_service import CommercialWorkflowService
from app.services.production_planning_service import ProductionPlanningService
from app.services.production_service import ProductionService


def get_production_planning_service() -> ProductionPlanningService:
    return ProductionPlanningService()


def get_production_service() -> ProductionService:
    planning_service = get_production_planning_service()
    return ProductionService(
        plan_repository=planning_service.plan_repository,
        planning_service=planning_service,
    )


def get_commercial_service() -> CommercialService:
    return CommercialService()


def get_commercial_workflow_service() -> CommercialWorkflowService:
    return CommercialWorkflowService()


def get_admin_service() -> AdminService:
    return AdminService()


def get_archive_service() -> ArchiveService:
    return ArchiveService()
