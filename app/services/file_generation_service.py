from __future__ import annotations

from contextlib import contextmanager
from typing import Iterator

from app.core.settings import get_settings
from app.domain.models.optimization_context import OptimizationContext
from app.domain.models.plate_order import PlateOrder
from app.services.optimization_service import OptimizationService
from core.commercial_offer import generate_commercial_offer_pdf, save_breakdown_to_excel
from core.commercial_offer_xlsx import generate_commercial_offer_xlsx
from core.config_and_data import PlateOrder as LegacyPlateOrder
from core.plates_preview_xlsx import build_plates_reconciliation_preview_xlsx
from core.visualization import visualize_plan


class FileGenerationService:
    def __init__(self) -> None:
        self.settings = get_settings()
        self.optimization_service = OptimizationService()

    @contextmanager
    def _legacy_order_context(self, order: PlateOrder) -> Iterator[None]:
        legacy_order = LegacyPlateOrder.from_dict(order.to_dict())
        legacy_order.apply_to_globals()
        yield

    def generate_preview_xlsx(
        self,
        *,
        output_path: str,
        plates_text: str,
        initial_user_plate_lines: list[str],
        forced_wide_line_indexes: list[int] | None = None,
    ) -> str:
        build_plates_reconciliation_preview_xlsx(
            output_path,
            plates_text=plates_text,
            initial_user_plate_lines=initial_user_plate_lines,
            forced_wide_line_indexes=forced_wide_line_indexes,
        )
        return output_path

    def generate_offer_pdf(self, *, order_data: list[dict], output_path: str, **kwargs) -> str:
        buffer = generate_commercial_offer_pdf(order_data, **kwargs)
        with open(output_path, "wb") as file:
            file.write(buffer.getvalue())
        return output_path

    def generate_offer_xlsx(self, *, order_data: list[dict], output_path: str, **kwargs) -> str:
        buffer = generate_commercial_offer_xlsx(order_data, **kwargs)
        with open(output_path, "wb") as file:
            file.write(buffer.getvalue())
        return output_path

    def save_breakdown(self, *, breakdown_tables: list[dict], output_path: str) -> str:
        save_breakdown_to_excel(breakdown_tables, output_path)
        return output_path

    def generate_visualization(self, *, order: PlateOrder, context: OptimizationContext, output_dir: str | None = None):
        with self._legacy_order_context(order):
            with self.optimization_service.legacy_runtime(context):
                return visualize_plan(output_dir or str(self.settings.outputs_dir))

