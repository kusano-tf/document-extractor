import xlrd

from .extractor_base import ExtractorBase


class XlsExtractor(ExtractorBase):
    def extract(self) -> str:
        wb = xlrd.open_workbook(str(self.path), on_demand=True)
        try:
            lines = []
            for sheet in wb.sheets():
                lines.append(f"[{sheet.name}]")
                for row_idx in range(sheet.nrows):
                    row = sheet.row_values(row_idx)
                    if all(str(cell_value).strip() == "" for cell_value in row):
                        continue
                    lines.append("\t".join(str(cell_value or "") for cell_value in row))
                lines.append("")
            return "\n".join(lines)
        finally:
            wb.release_resources()
