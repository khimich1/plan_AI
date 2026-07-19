from __future__ import annotations

from app.core.settings import get_settings
from app.domain.models.optimization_context import OptimizationContext
from app.domain.models.plate_order import PlateOrder
from app.services.optimization_service import OptimizationService
from core.commercial_offer import generate_commercial_offer_pdf, save_breakdown_to_excel
from core.commercial_offer_xlsx import generate_commercial_offer_xlsx
from core.plate_order_context import PlateOrderContext
from core.plates_preview_xlsx import build_plates_reconciliation_preview_xlsx
from core.ports.visualization import get_visualize_plan


class FileGenerationService:
    def __init__(self) -> None:
        self.settings = get_settings()
        self.optimization_service = OptimizationService()

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

    def generate_visualization(
        self,
        *,
        order: PlateOrder,
        context: OptimizationContext,
        ctx: PlateOrderContext,
        output_dir: str | None = None,
    ):
        ctx.hydrate_from_order(order)
        ctx.load_optimization_snapshot(
            optimization_result=context.optimization_result,
            plan_by_load=context.plan_by_load,
            load_to_reinforcement_map=context.load_to_reinforcement_map,
        )
        return get_visualize_plan()(
            output_dir or str(self.settings.outputs_dir),
            plate_order_ctx=ctx,
        )
