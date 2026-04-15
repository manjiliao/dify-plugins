import logging
import tempfile
from collections.abc import Generator
from pathlib import Path
from typing import Any, Optional

from dify_plugin import Tool
from dify_plugin.entities.tool import ToolInvokeMessage
from dify_plugin.file.file import File
from spire.xls import Workbook


EXCEL_MIME_TYPES = [
    "application/vnd.ms-excel",
    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    "application/vnd.ms-excel.sheet.macroEnabled.12",
    "application/vnd.ms-excel.sheet.binary.macroEnabled.12",
    "application/vnd.oasis.opendocument.spreadsheet",
    "application/octet-stream",
]

logger = logging.getLogger(__name__)


class ExcelToImageTool(Tool):
    def _invoke(
        self,
        tool_parameters: dict[str, Any],
        user_id: Optional[str] = None,
        conversation_id: Optional[str] = None,
        app_id: Optional[str] = None,
        message_id: Optional[str] = None,
    ) -> Generator[ToolInvokeMessage, None, None]:
        workbook = Workbook()
        temp_excel_path: Optional[Path] = None

        try:
            excel_content = tool_parameters.get("excel_content")
            if not isinstance(excel_content, File):
                raise ValueError("Invalid Excel content format. Expected a Dify File object.")

            image_format = self._normalize_image_format(tool_parameters.get("image_format"))
            remove_margins = self._to_bool(tool_parameters.get("remove_margins"), default=True)
            sheet_selector = str(tool_parameters.get("sheets") or "all").strip()
            scale_percent = self._parse_int(tool_parameters.get("scale_percent"), default=130, minimum=10, maximum=400)
            output_dpi = self._parse_int(tool_parameters.get("output_dpi"), default=150, minimum=72, maximum=600)

            filename = excel_content.filename or "workbook.xlsx"
            stem = Path(filename).stem or "workbook"
            request_context = {
                "file_name": filename,
                "image_format": image_format,
                "sheets": sheet_selector or "all",
                "remove_margins": remove_margins,
                "scale_percent": scale_percent,
                "output_dpi": output_dpi,
                "user_id": user_id,
                "conversation_id": conversation_id,
                "app_id": app_id,
                "message_id": message_id,
            }
            logger.info("excel_to_image invoke started: %s", request_context)

            with tempfile.TemporaryDirectory(prefix="excel_to_image_") as temp_dir:
                temp_dir_path = Path(temp_dir)
                temp_excel_path = temp_dir_path / filename
                temp_excel_path.write_bytes(excel_content.blob)

                workbook.LoadFromFile(str(temp_excel_path))
                workbook.ConverterSetting.XDpi = output_dpi
                workbook.ConverterSetting.YDpi = output_dpi
                workbook.ConverterSetting.ToImageWithoutMargins = remove_margins
                worksheet_total = workbook.Worksheets.Count
                selected_indexes = self._parse_sheet_selector(
                    workbook=workbook,
                    selector=sheet_selector,
                    worksheet_total=worksheet_total,
                )
                if not selected_indexes:
                    raise ValueError("No worksheets matched the provided selector.")

                exported_names: list[str] = []
                for sheet_index in selected_indexes:
                    sheet = workbook.Worksheets[sheet_index]
                    sheet.PageSetup.Zoom = scale_percent
                    if remove_margins:
                        self._strip_margins(sheet)

                    first_row = max(int(sheet.FirstRow), 1)
                    first_column = max(int(sheet.FirstColumn), 1)
                    last_row = max(int(sheet.LastRow), first_row)
                    last_column = max(int(sheet.LastColumn), first_column)

                    image = sheet.ToImage(first_row, first_column, last_row, last_column)
                    safe_sheet_name = self._sanitize_name(str(sheet.Name))
                    output_name = f"{stem}_{sheet_index + 1:02d}_{safe_sheet_name}.{image_format}"
                    output_path = temp_dir_path / output_name
                    image.Save(str(output_path))

                    mime_type = "image/png" if image_format == "png" else "image/jpeg"
                    exported_names.append(output_name)
                    yield self.create_blob_message(
                        blob=output_path.read_bytes(),
                        meta={"mime_type": mime_type, "file_name": output_name},
                    )

                selector_text = sheet_selector or "all"
                logger.info(
                    "excel_to_image invoke completed: file=%s exported=%s selector=%s format=%s scale=%s dpi=%s",
                    filename,
                    exported_names,
                    selector_text,
                    image_format,
                    scale_percent,
                    output_dpi,
                )
                yield self.create_text_message(
                    f"Converted {len(exported_names)} worksheet(s) from '{filename}' to {image_format.upper()} "
                    f"using selector '{selector_text}', scale {scale_percent}% and {output_dpi} DPI. "
                    f"Files: {', '.join(exported_names)}"
                )
        except Exception as exc:
            logger.exception("excel_to_image invoke failed")
            raise Exception(f"Failed to convert Excel to image: {exc}") from exc
        finally:
            workbook.Dispose()

    def _normalize_image_format(self, value: Any) -> str:
        normalized = str(value or "png").strip().lower()
        if normalized in {"jpeg", "jpg"}:
            return "jpg"
        if normalized != "png":
            raise ValueError("image_format must be png or jpg.")
        return normalized

    def _to_bool(self, value: Any, default: bool) -> bool:
        if value is None:
            return default
        if isinstance(value, bool):
            return value
        return str(value).strip().lower() in {"1", "true", "yes", "y", "on"}

    def _parse_int(self, value: Any, default: int, minimum: int, maximum: int) -> int:
        if value is None or str(value).strip() == "":
            return default
        try:
            number = int(str(value).strip())
        except ValueError as exc:
            raise ValueError(f"Value '{value}' must be an integer.") from exc
        if number < minimum or number > maximum:
            raise ValueError(f"Value '{number}' must be between {minimum} and {maximum}.")
        return number

    def _strip_margins(self, sheet: Any) -> None:
        sheet.PageSetup.LeftMargin = 0
        sheet.PageSetup.RightMargin = 0
        sheet.PageSetup.TopMargin = 0
        sheet.PageSetup.BottomMargin = 0

    def _parse_sheet_selector(self, workbook: Workbook, selector: str, worksheet_total: int) -> list[int]:
        if selector.lower() == "all":
            return list(range(worksheet_total))

        all_names = {str(workbook.Worksheets[index].Name).lower(): index for index in range(worksheet_total)}
        selected_indexes: list[int] = []

        for raw_token in selector.split(","):
            token = raw_token.strip()
            if not token:
                continue

            if "-" in token:
                range_parts = [part.strip() for part in token.split("-", 1)]
                if len(range_parts) != 2 or not all(part.isdigit() for part in range_parts):
                    raise ValueError(f"Invalid sheet range: {token}")
                start = int(range_parts[0])
                end = int(range_parts[1])
                if start < 1 or end < 1 or start > worksheet_total or end > worksheet_total:
                    raise ValueError(f"Sheet range out of bounds: {token}")
                step = 1 if start <= end else -1
                for number in range(start, end + step, step):
                    selected_indexes.append(number - 1)
                continue

            if token.isdigit():
                number = int(token)
                if number < 1 or number > worksheet_total:
                    raise ValueError(f"Sheet index out of bounds: {token}")
                selected_indexes.append(number - 1)
                continue

            lowered = token.lower()
            if lowered not in all_names:
                raise ValueError(f"Sheet name not found: {token}")
            selected_indexes.append(all_names[lowered])

        deduplicated: list[int] = []
        for index in selected_indexes:
            if index not in deduplicated:
                deduplicated.append(index)
        return deduplicated

    def _sanitize_name(self, value: str) -> str:
        clean = "".join(ch if ch.isalnum() or ch in {"-", "_"} else "_" for ch in value).strip("_")
        return clean or "sheet"
