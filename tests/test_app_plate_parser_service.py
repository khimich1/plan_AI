from app.services.plate_parser_service import PlateParserService


def test_parse_plate_text_returns_order_without_globals_dependency():
    service = PlateParserService()

    result = service.parse_plate_text("ПБ 78-12-8п 2\n0,32x6,63 - 4")

    assert result.order.plate_load_details
    assert not result.unparsed_lines
    assert result.order.plates_1_2 == [7.8, 7.8]
    assert result.order.plates_0_32 == [6.63, 6.63, 6.63, 6.63]
    assert len(result.order.plates_1_2) + len(result.order.plates_0_32) == 6


def test_parse_plate_text_bare_format_without_pb_prefix():
    service = PlateParserService()
    text = "\n".join(
        [
            "71-12-8  3",
            "71-9,20-8   1",
            "65,6-12-12,5  2",
            "32,5-6,65-8  1",
            "31,5-10,8-8  3",
        ]
    )

    result = service.parse_plate_text(text)

    assert not result.unparsed_lines
    assert result.order.plate_load_details
    assert sum(result.order.plate_load_details.values()) == 10

