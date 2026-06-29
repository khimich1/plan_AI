"""Re-export KP save orchestration from core (A2 offers slice)."""

from core.kp_persistence_service import KpPersistenceService

__all__ = ["KpPersistenceService"]
