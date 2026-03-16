"""Excel report generation using formulas for all derived values."""

from datetime import date, time
from pathlib import Path
from typing import Dict, Iterable, List

from openpyxl import Workbook
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.worksheet import Worksheet

from .models import DayBounds, EmployeeCalendar

HEADER_FILL = PatternFill(fill_type="solid", start_color="D9E1F2", end_color="D9E1F2")
LIGHT_RED_FILL = PatternFill(fill_type="solid", start_color="FFC7CE", end_color="FFC7CE")
DARK_RED_FILL = PatternFill(fill_type="solid", start_color="9C0006", end_color="9C0006")
LIGHT_GREEN_FILL = PatternFill(fill_type="solid", start_color="C6EFCE", end_color="C6EFCE")
DARK_GREEN_FILL = PatternFill(fill_type="solid", start_color="006100", end_color="006100")
BLACK_FONT = Font(color="000000")
THIN_BORDER = Border(
    left=Side(style="thin"),
    right=Side(style="thin"),
    top=Side(style="thin"),
    bottom=Side(style="thin"),
)
THICK_SIDE = Side(style="thick")

def write_report(
    output_path: Path,
    calendar: EmployeeCalendar,
    default_official_time: time = time(9, 0),
) -> None:
    default_leave_time: time = time(18, 0)

    """Create report workbook from aggregated employee daily attendance bounds."""
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Отчет"

    all_dates = _collect_sorted_dates(calendar)
    _write_header(sheet, all_dates)
    _write_body(sheet, calendar, all_dates, default_official_time, default_leave_time)
    _apply_layout(sheet, all_dates)

    output_path.parent.mkdir(parents=True, exist_ok=True)
    workbook.save(output_path)


def _collect_sorted_dates(calendar: EmployeeCalendar) -> List[date]:
    """Collect and sort all dates present in the calendar."""
    day_set = set()
    for employee_days in calendar.values():
        day_set.update(employee_days.keys())
    return sorted(day_set)


def _write_header(sheet: Worksheet, all_dates: List[date]) -> None:
    """Write the first row with fixed headers, date columns and average columns."""
    sheet.cell(row=1, column=1, value="ФИО")
    sheet.cell(row=1, column=2, value="Начало рабочего дня")
    sheet.cell(row=1, column=3, value="Конец рабочего дня")
    sheet.cell(row=1, column=4, value="Продолжительность работы")
    sheet.cell(row=1, column=5, value="Показатель")

    first_date_column = 6
    for date_index, day_value in enumerate(all_dates):
        column_index = first_date_column + date_index
        cell = sheet.cell(row=1, column=column_index, value=day_value)
        cell.number_format = "dd.mm.yyyy"

    average_headers = (
        "Среднее время прихода",
        "Среднее время ухода",
        "Среднее время работы",
        "Среднее отклонение прихода",
        "Среднее время переработок",
        "Среднее отсутствие в течение дня",
        "Средняя длительность факт (с учетом отсутствия в течение дня)",
        "Переработка факт (с учетом отсутствия в течение дня)",
    )
    first_average_column = first_date_column + len(all_dates)
    for offset, label in enumerate(average_headers):
        sheet.cell(row=1, column=first_average_column + offset, value=label)


def _write_body(
    sheet: Worksheet,
    calendar: EmployeeCalendar,
    all_dates: List[date],
    default_official_time: time,
    default_leave_time: time
) -> None:
    """Write employee rows with values and formulas."""
    first_data_column = 6
    first_average_column = first_data_column + len(all_dates)

    for employee_index, employee_name in enumerate(sorted(calendar)):
        arrival_row = 2 + employee_index * 8
        leave_row = arrival_row + 1
        work_row = arrival_row + 2
        delta_row = arrival_row + 3
        overtime_row = arrival_row + 4
        absence_row = arrival_row + 5
        work_minus_absence_row = arrival_row + 6
        overtime_minus_absence_row = arrival_row + 7

        sheet.cell(row=arrival_row, column=1, value=employee_name)

        sheet.cell(row=arrival_row, column=2, value=default_official_time)
        sheet.cell(row=arrival_row, column=2).number_format = "hh:mm"

        sheet.cell(row=arrival_row, column=4, value=time(9, 0))
        sheet.cell(row=arrival_row, column=4).number_format = "hh:mm"

        # Конец рабочего дня = начало рабочего дня + продолжительность работы
        end_of_day_formula = f"=B{arrival_row}+D{arrival_row}"
        sheet.cell(row=arrival_row, column=3, value=end_of_day_formula)
        sheet.cell(row=arrival_row, column=3).number_format = "hh:mm"

        sheet.cell(row=arrival_row, column=5, value="Время прихода")
        sheet.cell(row=leave_row, column=5, value="Время ухода")
        sheet.cell(row=work_row, column=5, value="Длительность факт")
        sheet.cell(row=delta_row, column=5, value="Отклонение по времени прихода")
        sheet.cell(row=overtime_row, column=5, value="Переработка")
        sheet.cell(row=absence_row, column=5, value="Отсутствие в течение дня")
        sheet.cell(
            row=work_minus_absence_row,
            column=5,
            value="Длительность факт (с учетом отсутствия в течение дня)",
        )
        sheet.cell(
            row=overtime_minus_absence_row,
            column=5,
            value="Переработка факт (с учетом отсутствия в течение дня)",
        )

        employee_days: Dict[date, DayBounds] = calendar[employee_name]
        for day_offset, day_value in enumerate(all_dates):
            column_index = first_data_column + day_offset
            column_letter = get_column_letter(column_index)

            day_bounds = employee_days.get(day_value)
            if day_bounds is not None:
                arrival_cell = sheet.cell(
                    row=arrival_row,
                    column=column_index,
                    value=day_bounds.arrival_time,
                )
                leave_cell = sheet.cell(
                    row=leave_row,
                    column=column_index,
                    value=day_bounds.departure_time,
                )
                absence_cell = sheet.cell(
                    row=absence_row,
                    column=column_index,
                    value=day_bounds.absence_duration,
                )
                arrival_cell.number_format = "hh:mm"
                leave_cell.number_format = "hh:mm"
                absence_cell.number_format = "[h]:mm"

            work_formula = (
                f'=IF(OR({column_letter}{arrival_row}="",{column_letter}{leave_row}=""),"",'
                f"{column_letter}{leave_row}-{column_letter}{arrival_row})"
            )
            delta_formula = (
                f'=IF({column_letter}{arrival_row}="","",'
                f'IF({column_letter}{arrival_row}>=$B${arrival_row},'
                f'TEXT({column_letter}{arrival_row}-$B${arrival_row},"ч:мм"),'
                f'TEXT($B${arrival_row}-{column_letter}{arrival_row},"-ч:мм")))'
            )
            overtime_formula = (
                f'=IF({column_letter}{work_row}="","",'
                f'IF({column_letter}{work_row}>=$D${arrival_row},'
                f'TEXT({column_letter}{work_row}-$D${arrival_row},"ч:мм"),'
                f'TEXT($D${arrival_row}-{column_letter}{work_row},"-ч:мм")))'
            )
            work_minus_absence_formula = (
                f'=IF(OR({column_letter}{work_row}="",{column_letter}{absence_row}=""),"",'
                f"{column_letter}{work_row}-{column_letter}{absence_row})"
            )
            # Переработка факт (с учетом отсутствия): переработка считается
            # от фактической длительности с учетом отсутствия, а не как
            # разность текстовой переработки и отсутствия.
            overtime_minus_absence_formula = (
                f'=IF({column_letter}{work_minus_absence_row}="","",'
                f'IF({column_letter}{work_minus_absence_row}>=$D${arrival_row},'
                f'TEXT({column_letter}{work_minus_absence_row}-$D${arrival_row},"ч:мм"),'
                f'TEXT($D${arrival_row}-{column_letter}{work_minus_absence_row},"-ч:мм")))'
            )


            work_cell = sheet.cell(
                row=work_row, column=column_index, value=work_formula
            )
            delta_cell = sheet.cell(
                row=delta_row, column=column_index, value=delta_formula
            )
            work_minus_absence_cell = sheet.cell(
                row=work_minus_absence_row,
                column=column_index,
                value=work_minus_absence_formula,
            )
            overtime_minus_absence_cell = sheet.cell(
                row=overtime_minus_absence_row,
                column=column_index,
                value=overtime_minus_absence_formula,
            )
            work_cell.number_format = "[h]:mm"
            delta_cell.number_format = "@"
            sheet.cell(
                row=overtime_row, column=column_index, value=overtime_formula
            ).number_format = "[h]:mm"
            work_minus_absence_cell.number_format = "[h]:mm"
            overtime_minus_absence_cell.number_format = "[h]:mm"

        if all_dates:
            start_column_letter = get_column_letter(first_data_column)
            end_column_letter = get_column_letter(first_data_column + len(all_dates) - 1)

            avg_arrival_formula = (
                f'=IFERROR(AVERAGE({start_column_letter}{arrival_row}:{end_column_letter}{arrival_row}),"")'
            )
            avg_leave_formula = (
                f'=IFERROR(AVERAGE({start_column_letter}{leave_row}:{end_column_letter}{leave_row}),"")'
            )
            avg_work_formula = (
                f'=IFERROR(AVERAGE({start_column_letter}{work_row}:{end_column_letter}{work_row}),"")'
            )
            avg_delta_formula = (
                f'=IFERROR(IF(AVERAGE({start_column_letter}{arrival_row}:{end_column_letter}{arrival_row})'
                f'>=$B${arrival_row},'
                f'TEXT(AVERAGE({start_column_letter}{arrival_row}:{end_column_letter}{arrival_row})'
                f'-$B${arrival_row},"ч:мм"),'
                f'TEXT($B${arrival_row}-AVERAGE({start_column_letter}{arrival_row}:'
                f'{end_column_letter}{arrival_row}),"-ч:мм")),"")'
            )
            avg_overtime_formula = (
                f'=IFERROR(IF(AVERAGE({start_column_letter}{work_row}:{end_column_letter}{work_row})'
                f'>=$D${arrival_row},'
                f'TEXT(AVERAGE({start_column_letter}{work_row}:{end_column_letter}{work_row})'
                f'-$D${arrival_row},"ч:мм"),'
                f'TEXT($D${arrival_row}-AVERAGE({start_column_letter}{work_row}:'
                f'{end_column_letter}{work_row}),"-ч:мм")),"")'
            )
            avg_absence_formula = (
                f'=IFERROR(AVERAGE({start_column_letter}{absence_row}:{end_column_letter}{absence_row}),"")'
            )
            avg_work_minus_absence_formula = (
                f'=IFERROR(AVERAGE({start_column_letter}{work_minus_absence_row}:{end_column_letter}{work_minus_absence_row}),"")'
            )
            # Средняя «переработка факт (с учетом отсутствия)» считается
            # по среднему значению длительности с учетом отсутствия.
            avg_overtime_minus_absence_formula = (
                f'=IFERROR(IF(AVERAGE({start_column_letter}{work_minus_absence_row}:{end_column_letter}{work_minus_absence_row})'
                f'>=$D${arrival_row},'
                f'TEXT(AVERAGE({start_column_letter}{work_minus_absence_row}:{end_column_letter}{work_minus_absence_row})'
                f'-$D${arrival_row},"ч:мм"),'
                f'TEXT($D${arrival_row}-AVERAGE({start_column_letter}{work_minus_absence_row}:'
                f'{end_column_letter}{work_minus_absence_row}),"-ч:мм")),"")'
            )

            avg_arrival_cell = sheet.cell(
                row=arrival_row, column=first_average_column, value=avg_arrival_formula
            )
            avg_leave_cell = sheet.cell(
                row=leave_row, column=first_average_column + 1, value=avg_leave_formula
            )
            avg_work_cell = sheet.cell(
                row=work_row, column=first_average_column + 2, value=avg_work_formula
            )
            avg_delta_cell = sheet.cell(
                row=delta_row, column=first_average_column + 3, value=avg_delta_formula
            )
            avg_overtime_cell = sheet.cell(
                row=overtime_row,
                column=first_average_column + 4,
                value=avg_overtime_formula,
            )
            avg_absence_cell = sheet.cell(
                row=absence_row,
                column=first_average_column + 5,
                value=avg_absence_formula,
            )
            avg_work_minus_absence_cell = sheet.cell(
                row=work_minus_absence_row,
                column=first_average_column + 6,
                value=avg_work_minus_absence_formula,
            )
            avg_overtime_minus_absence_cell = sheet.cell(
                row=overtime_minus_absence_row,
                column=first_average_column + 7,
                value=avg_overtime_minus_absence_formula,
            )

            avg_arrival_cell.number_format = "hh:mm"
            avg_leave_cell.number_format = "hh:mm"
            avg_work_cell.number_format = "[h]:mm"
            avg_delta_cell.number_format = "@"
            avg_overtime_cell.number_format = "[h]:mm"
            avg_absence_cell.number_format = "[h]:mm"
            avg_work_minus_absence_cell.number_format = "[h]:mm"
            avg_overtime_minus_absence_cell.number_format = "[h]:mm"


def _apply_layout(sheet: Worksheet, all_dates: List[date]) -> None:
    """Apply basic workbook style, widths and frozen panes."""
    total_columns = 4 + len(all_dates) + 9
    last_row = max(1, sheet.max_row)

    sheet.freeze_panes = "F2"
    sheet.column_dimensions["A"].width = 34
    sheet.column_dimensions["B"].width = 18
    sheet.column_dimensions["C"].width = 28
    sheet.column_dimensions["E"].width = 60

    for column_index in range(4, total_columns + 1):
        letter = get_column_letter(column_index)
        sheet.column_dimensions[letter].width = 14

    for column_index in range(1, total_columns + 1):
        header_cell = sheet.cell(row=1, column=column_index)
        header_cell.font = Font(bold=True)
        header_cell.fill = HEADER_FILL
        header_cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    for row_index in range(1, last_row + 1):
        for column_index in range(1, total_columns + 1):
            cell = sheet.cell(row=row_index, column=column_index)
            cell.border = THIN_BORDER
            if row_index > 1 and column_index >= 4:
                cell.alignment = Alignment(horizontal="center", vertical="center")

    # Толстые рамки вокруг блока строк одного сотрудника.
    # Определяем начало блока по строкам, где в колонке E стоит "Время прихода".
    if last_row > 1:
        employee_starts = []
        for row_index in range(2, last_row + 1):
            marker = sheet.cell(row=row_index, column=5).value
            if marker == "Время прихода":
                employee_starts.append(row_index)

        for idx, start_row in enumerate(employee_starts):
            end_row = (
                employee_starts[idx + 1] - 1
                if idx + 1 < len(employee_starts)
                else last_row
            )

            for column_index in range(1, total_columns + 1):
                cell = sheet.cell(row=start_row, column=column_index)
                cell.border = Border(
                    left=cell.border.left,
                    right=cell.border.right,
                    top=THICK_SIDE,
                    bottom=cell.border.bottom,
                )

                cell = sheet.cell(row=end_row, column=column_index)
                cell.border = Border(
                    left=cell.border.left,
                    right=cell.border.right,
                    top=cell.border.top,
                    bottom=THICK_SIDE,
                )

            for row_index in range(start_row, end_row + 1):
                cell = sheet.cell(row=row_index, column=1)
                cell.border = Border(
                    left=THICK_SIDE,
                    right=cell.border.right,
                    top=cell.border.top,
                    bottom=cell.border.bottom,
                )

                cell = sheet.cell(row=row_index, column=total_columns)
                cell.border = Border(
                    left=cell.border.left,
                    right=THICK_SIDE,
                    top=cell.border.top,
                    bottom=cell.border.bottom,
                )

    _apply_conditional_formatting(sheet, all_dates, last_row)


def _apply_conditional_formatting(
    sheet: Worksheet, all_dates: List[date], last_row: int
) -> None:
    """Apply conditional formatting for delta and overtime rows only."""
    if not all_dates:
        return
    first_data_column = 6
    start_col = get_column_letter(first_data_column)
    end_col = get_column_letter(first_data_column + len(all_dates) - 1)

    # Определяем начало блоков сотрудников по строкам с "Время прихода".
    employee_starts = []
    for row_index in range(2, last_row + 1):
        marker = sheet.cell(row=row_index, column=5).value
        if marker == "Время прихода":
            employee_starts.append(row_index)

    for employee_index, arrival_row in enumerate(employee_starts):
        delta_row = arrival_row + 3
        overtime_row = arrival_row + 4

        # Отклонение по времени прихода
        delta_range = f"{start_col}{delta_row}:{end_col}{delta_row}"
        # diff = фактическое время прихода - плановое время
        # зелёный: diff < -15 минут
        delta_green_rule = FormulaRule(
            formula=[
                f'AND({start_col}{arrival_row}<>"",'
                f'{start_col}{arrival_row}-$B${arrival_row}<-1/96)'
            ],
            fill=LIGHT_GREEN_FILL,
            font=BLACK_FONT,
        )
        # красный: diff > 15 минут
        delta_red_rule = FormulaRule(
            formula=[
                f'AND({start_col}{arrival_row}<>"",'
                f'{start_col}{arrival_row}-$B${arrival_row}>1/96)'
            ],
            fill=LIGHT_RED_FILL,
            font=BLACK_FONT,
        )
        sheet.conditional_formatting.add(delta_range, delta_green_rule)
        sheet.conditional_formatting.add(delta_range, delta_red_rule)

        # Переработка
        work_row = arrival_row + 2
        overtime_range = f"{start_col}{overtime_row}:{end_col}{overtime_row}"
        # diff = фактическая длительность - плановая длительность
        diff_expr = f"{start_col}{work_row}-$D${arrival_row}"

        # 0 < diff <= 30 минут  -> светло-зелёный (игнорируем почти нулевые значения)
        overtime_light_green = FormulaRule(
            formula=[
                f'AND({start_col}{work_row}<>"",'
                f'{diff_expr}>0,'
                f'{diff_expr}<=1/48,'
                f'ABS({diff_expr})>=1/1440)'
            ],
            fill=LIGHT_GREEN_FILL,
            font=BLACK_FONT,
        )
        # diff > 30 минут -> тёмно-зелёный
        overtime_dark_green = FormulaRule(
            formula=[
                f'AND({start_col}{work_row}<>"",'
                f'{diff_expr}>1/48,'
                f'ABS({diff_expr})>=1/1440)'
            ],
            fill=DARK_GREEN_FILL,
            font=BLACK_FONT,
        )
        # -30 минут <= diff < 0 -> светло-красный
        overtime_light_red = FormulaRule(
            formula=[
                f'AND({start_col}{work_row}<>"",'
                f'{diff_expr}<0,'
                f'{diff_expr}>=-1/48,'
                f'ABS({diff_expr})>=1/1440)'
            ],
            fill=LIGHT_RED_FILL,
            font=BLACK_FONT,
        )
        # diff < -30 минут -> тёмно-красный
        overtime_dark_red = FormulaRule(
            formula=[
                f'AND({start_col}{work_row}<>"",'
                f'{diff_expr}<-1/48,'
                f'ABS({diff_expr})>=1/1440)'
            ],
            fill=DARK_RED_FILL,
            font=BLACK_FONT,
        )

        sheet.conditional_formatting.add(overtime_range, overtime_light_green)
        sheet.conditional_formatting.add(overtime_range, overtime_dark_green)
        sheet.conditional_formatting.add(overtime_range, overtime_light_red)
        sheet.conditional_formatting.add(overtime_range, overtime_dark_red)

        # Переработка факт (с учетом отсутствия в течение дня)
        work_minus_absence_row = arrival_row + 6
        overtime_minus_absence_row = arrival_row + 7
        overtime_factual_range = (
            f"{start_col}{overtime_minus_absence_row}:{end_col}{overtime_minus_absence_row}"
        )
        diff_factual_expr = f"{start_col}{work_minus_absence_row}-$D${arrival_row}"

        overtime_factual_light_green = FormulaRule(
            formula=[
                f'AND({start_col}{work_minus_absence_row}<>"",'
                f'{diff_factual_expr}>0,'
                f'{diff_factual_expr}<=1/48,'
                f'ABS({diff_factual_expr})>=1/1440)'
            ],
            fill=LIGHT_GREEN_FILL,
            font=BLACK_FONT,
        )
        overtime_factual_dark_green = FormulaRule(
            formula=[
                f'AND({start_col}{work_minus_absence_row}<>"",'
                f'{diff_factual_expr}>1/48,'
                f'ABS({diff_factual_expr})>=1/1440)'
            ],
            fill=DARK_GREEN_FILL,
            font=BLACK_FONT,
        )
        overtime_factual_light_red = FormulaRule(
            formula=[
                f'AND({start_col}{work_minus_absence_row}<>"",'
                f'{diff_factual_expr}<0,'
                f'{diff_factual_expr}>=-1/48,'
                f'ABS({diff_factual_expr})>=1/1440)'
            ],
            fill=LIGHT_RED_FILL,
            font=BLACK_FONT,
        )
        overtime_factual_dark_red = FormulaRule(
            formula=[
                f'AND({start_col}{work_minus_absence_row}<>"",'
                f'{diff_factual_expr}<-1/48,'
                f'ABS({diff_factual_expr})>=1/1440)'
            ],
            fill=DARK_RED_FILL,
            font=BLACK_FONT,
        )

        sheet.conditional_formatting.add(
            overtime_factual_range, overtime_factual_light_green
        )
        sheet.conditional_formatting.add(
            overtime_factual_range, overtime_factual_dark_green
        )
        sheet.conditional_formatting.add(
            overtime_factual_range, overtime_factual_light_red
        )
        sheet.conditional_formatting.add(
            overtime_factual_range, overtime_factual_dark_red
        )