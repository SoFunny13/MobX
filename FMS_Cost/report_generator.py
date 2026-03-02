from io import BytesIO
import openpyxl
from openpyxl.styles import Font, Alignment


def generate_report(date_str, exchange_rate, stats_rows, offer_id_map, source_id_map,
                    no_convert_sources=None):
    """Generate .xlsx report and return as bytes.

    Args:
        date_str: Date string for the report.
        exchange_rate: USD exchange rate (RUB per 1 USD).
        stats_rows: List of dicts with keys: offer_name, source_name, cost_rub.
        offer_id_map: Dict offer_name -> offer_id (int).
        source_id_map: Dict source_name -> source_id (int).
        no_convert_sources: Set of source names to keep in RUB (no USD conversion).

    Returns:
        bytes: Excel file content.
    """
    if no_convert_sources is None:
        no_convert_sources = set()

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Отчет"

    # Data rows (no header)
    for row in stats_rows:
        offer_id = offer_id_map.get(row["offer_name"], "")
        source_id = source_id_map.get(row["source_name"], "")
        cost_rub = float(row["cost_rub"]) if row["cost_rub"] else 0.0
        if row["source_name"] in no_convert_sources:
            cost = round(cost_rub)
        else:
            cost = round(cost_rub / exchange_rate) if exchange_rate else 0
        ws.append([date_str, offer_id, source_id, cost])

    # Column widths
    ws.column_dimensions["A"].width = 14
    ws.column_dimensions["B"].width = 12
    ws.column_dimensions["C"].width = 16
    ws.column_dimensions["D"].width = 16

    buf = BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.getvalue()
