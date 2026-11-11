# --------------------------- DOCUMENTATION ---------------------------#
# --------------------------- IMPORTS ---------------------------#
import base64
import os
import time
from typing import Any, cast

import pandas as pd
import xlsxwriter
import xlwings as xw
from loguru import logger


# --------------------------- CONSTANTS ---------------------------#
# --------------------------- FUNCTIONS ---------------------------#
def data_from_excel_sheet(excel_filepath: str, read_map: dict[str, list[str]]):
    """
    Reads values from cells across one or more sheets in a single session using xlwings(requires Excel to be installed in the system)\n
    **read_map:** {sheetname: [cell_address, ...]}\n
    **Returns:** {sheetname: {cell_address: value}} / {cell_address: value}
    """
    result: dict[str, dict[str, Any]] = {}
    with xw.App(visible=False) as app:
        src = app.books.open(excel_filepath)
        for sheet, cells in read_map.items():
            ws = src.sheets[sheet]
            result[sheet] = {cell: ws.range(cell).value for cell in cells}
        src.close()
    return next(iter(result.values())) if len(result) == 1 else result


def data_to_excel_sheet(
    data_map: dict[str, dict[str, Any]],
    excel_filepath: str,
    header: bool = True,
    vertical_list: bool = False,
    table_fmt: bool = False,
    table_total: bool = False,
):
    """
    Updates excel sheet (creates if not existing) with data (w/o index) using xlwings(requires Excel to be installed in the system)\n
    **data_map:** {sheetname: {cell_address/range: data variable}}\n
    **vertical_list:** True if passing a vertical list of values
    **table_fmt:** format data as table
    """
    with xw.App(visible=False) as app:
        new_dst = False
        if not os.path.exists(excel_filepath):
            dst = app.books.add()
            dst.save(excel_filepath)
            new_dst = True
        else:
            dst = app.books.open(excel_filepath)

        for sheet, address in data_map.items():
            try:
                ws = dst.sheets[sheet]
            except:
                ws = dst.sheets.add(sheet, after=dst.sheets[-1])
            for cell, data in address.items():
                start = ws.range(cell)
                if type(data) == pd.DataFrame:
                    try:
                        if start.value is not None:
                            start.expand().clear_contents()
                    except Exception:
                        pass  # no data, fine to proceed writing

                    start.options(index=False, header=header).value = data
                elif isinstance(data, (list, tuple)):
                    for i, val in enumerate(data):
                        if val == None:
                            continue
                        if vertical_list:
                            start.offset(row_offset=i).value = val
                        else:
                            start.offset(column_offset=i).value = val
                else:
                    start.value = data

                if table_fmt:
                    data_range = start.expand("table")
                    table = ws.api.ListObjects.Add(
                        SourceType=1,
                        Source=data_range.api,
                        XlListObjectHasHeaders=1 if header else 0,
                    )
                    # table.Name = name # cannot take None
                    table.TableStyle = "TableStyleMedium2"
                    table.ShowTotals = table_total

            ws.autofit()
            dst.save()
            logger.info(f"{sheet} sheet from {excel_filepath} successfully updated")

        if new_dst:
            dst.sheets["Sheet1"].delete()  # remove default sheet
        dst.close()


def copy_excel_sheet(
    src_path: str, dst_path: str, sheet_names: list[str] | None = None, first_pos: bool = False
):
    """
    Copy each sheet / selected sheets from source to destination file at the end(default) using xlwings(requires Excel to be installed in the system).
    Creates file if dst file does not exist
    """
    with xw.App(visible=False) as app:
        try:
            src = app.books.open(src_path)
        except FileNotFoundError:
            logger.error(f"{src_path} not found")
            raise

        new_dst = False
        if not os.path.exists(dst_path):
            dst = app.books.add()
            dst.save(dst_path)
            new_dst = True
        else:
            dst = app.books.open(dst_path)

        # preserving internal cross-references between them (for loop doesn't)
        sheet_list = sheet_names if sheet_names else [s.name for s in src.sheets]
        src.api.Worksheets(sheet_list).Copy(
            Before=dst.sheets[0].api
        ) if first_pos else src.api.Worksheets(sheet_list).Copy(After=dst.sheets[-1].api)

        if new_dst:
            dst.sheets[0].delete()  # remove default sheet

        src.close()
        dst.save()
        dst.close()


def hide_excel_sheet(excel_filepath: str, sheet_names: list[str] | str):
    """
    Hides the given sheets in the workbook using xlwings(requires Excel to be installed in the system)
    """
    with xw.App(visible=False) as app:
        wb = app.books.open(excel_filepath)
        if isinstance(sheet_names, str):
            wb.sheets[sheet_names].api.Visible = False
        else:
            for sheet in sheet_names:
                wb.sheets[sheet].api.Visible = False

        wb.save()
        wb.close()


def img_from_excel_sheet(
    excel_filepath: str,
    read_map: dict[str, dict[str, str | None]],
    output_dir: str,
    save_img: bool = True,
):
    """
    Converts range to image using xlwings(requires Excel to be installed in the system)\n
    **read_map:** {sheetname: {cell_range:img_name, ...]}   If img_name is None, outputs cell range as name\n
    RETURNS base64 encoded images to be used in HTML email body
    """
    output_dir = os.path.abspath(output_dir)
    with xw.App(visible=True) as app:
        src = app.books.open(excel_filepath)
        # out_paths: list[str] = []
        out_imgs: list[str] = []
        for sheet, ranges in read_map.items():
            ws = src.sheets[sheet]
            for cell_range, img_name in ranges.items():
                rng = ws.range(cell_range)

                # Make sure Excel actually renders the sheet/range
                ws.activate()
                time.sleep(2)  # 2 for remote location files

                # need to pass args as multiple xw objs in program mess up
                rng.api.CopyPicture(Appearance=1, Format=2)  # xlScreen=1, xlBitmap=2
                # Paste into a new chart object, then export as an actual image file
                chart_obj = ws.api.ChartObjects().Add(0, 0, rng.width, rng.height)
                chart = chart_obj.Chart
                chart.Paste()

                output_path = (
                    os.path.join(output_dir, f"{img_name}.png")
                    if img_name
                    else os.path.join(output_dir, f"{sheet}_{cell_range.replace(':', '_')}.png")
                )
                chart.Export(output_path)
                # out_paths.append(output_path)
                chart_obj.Delete()

                # Read the image and encode to base64
                with open(output_path, "rb") as img_file:
                    b64 = base64.b64encode(img_file.read()).decode("utf-8")
                out_imgs.append(f"data:image/png;base64,{b64}")

                if not save_img:
                    os.remove(output_path)

        src.close()  # close without saving, so the temp chart never touches the saved file
    return out_imgs


def pdf_from_excel_sheet(
    excel_filepath: str, read_map: dict[str, dict[str, str | None]], output_dir: str
):
    """
    Converts range to pdf using xlwings(requires Excel to be installed in the system)\n
    **read_map:** {sheetname: {cell_range:pdf_name, ...]}   If pdf_name is None, outputs cell range as name\n
    """
    output_dir = os.path.abspath(output_dir)
    with xw.App(visible=False) as app:
        src = app.books.open(excel_filepath)

        for sheet, ranges in read_map.items():
            ws = src.sheets[sheet]
            for cell_range, pdf_name in ranges.items():
                rng = ws.range(cell_range)
                output_path = (
                    os.path.join(output_dir, f"{pdf_name}.pdf")
                    if pdf_name
                    else os.path.join(output_dir, f"{sheet}_{cell_range.replace(':', '_')}.pdf")
                )

                rng.api.ExportAsFixedFormat(0, output_path)  # xlTypePDF=0

        src.close()  # close without saving, so the temp chart never touches the saved file


def format_table_sheet(
    df: pd.DataFrame,
    writer: pd.ExcelWriter,
    sheet_name: str,
    table_name: str | None = None,
    curr_cols: list[str] | None = None,
    need_totals: bool = False,
    cols_fmt: list | None = None,
):
    """
    Formats DF data as a table with a new sheet using xlsxwriter\n
    **curr_cols:** Formats column as currency. DEFAULT VALUE: "revenue"\n
    Call the fn in: with pd.ExcelWriter(file_path, engine="xlsxwriter") as writer:
    """
    workbook = cast(xlsxwriter.Workbook, writer.book)
    worksheet = writer.sheets[sheet_name]
    rows, cols = df.shape  # Get table dimensions

    last_row = rows + 1 if need_totals else rows  # reserves totals row

    # # Add Excel table (includes header row automatically)
    worksheet.add_table(
        0,
        0,
        last_row,
        cols - 1,
        {
            "name": sheet_name if not table_name else table_name,
            "columns": cols_fmt if cols_fmt else [({"header": col}) for col in df.columns],
            "style": "Table Style Medium 2",
            "total_row": need_totals,
        },
    )

    # Currency format for revenue
    currency_fmt = workbook.add_format(
        {"num_format": '_($* #,##0.00_);_($* (#,##0.00);_($* "-"??_);_(@_)'}
    )

    # Find revenue column index
    curr_cols = ["revenue"] if curr_cols is None else curr_cols
    for col in curr_cols:
        if col in df.columns.str.strip().str.lower():
            idx = df.columns.str.strip().str.lower().get_loc(col)

            # Apply format ONLY to that column
            worksheet.set_column(idx, idx, None, currency_fmt)

    # Auto-fit columns
    worksheet.autofit()


# --------------------------- MAIN CODE ---------------------------#
if __name__ == "__main__":
    pass
