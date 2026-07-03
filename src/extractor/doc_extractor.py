import textract  # type: ignore[import]

from .extractor_base import ExtractorBase


class DocExtractor(ExtractorBase):
    def extract(self) -> str:
        try:
            text = textract.process(str(self.path), extension="doc")
        except Exception as exc:
            raise RuntimeError(f"Failed to extract .doc file: {exc}") from exc

        if isinstance(text, bytes):
            text = text.decode("utf-8", errors="replace")
        return text.strip()
