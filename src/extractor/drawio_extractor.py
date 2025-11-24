import re
import xml.etree.ElementTree as ET

from .extractor_base import ExtractorBase


class DrawioExtractor(ExtractorBase):
    """Draw.io ファイルからテキストを抽出するクラス"""

    def extract(self) -> str:
        """draw.io ファイルからすべてのテキストを抽出"""
        try:
            tree = ET.parse(self.path)
            root = tree.getroot()
        except ET.ParseError as e:
            raise ValueError(f"Failed to parse drawio file: {e}")

        texts = self._extract_texts_from_root(root)
        return "\n".join(texts)

    def _extract_texts_from_root(self, root: ET.Element) -> list[str]:
        """ルート要素からテキストを再帰的に抽出"""
        texts = []

        # mxCell 要素からテキストを抽出
        for cell in root.findall(".//mxCell"):
            text = self._extract_text_from_cell(cell)
            if text:
                texts.append(text)

        return texts

    def _extract_text_from_cell(self, cell: ET.Element) -> str:
        """mxCell 要素からテキストを抽出"""
        # value 属性をチェック
        value = cell.get("value")
        if value and value.strip():
            # HTML タグを削除
            text = self._remove_html_tags(value.strip())
            return text

        # value 属性がない場合、mxGeometry内を探索
        geometry = cell.find("mxGeometry")
        if geometry is not None:
            # Richtext の場合も対応
            # draw.io では value 属性が主だが、念のため確認
            pass

        return ""

    def _remove_html_tags(self, text: str) -> str:
        """テキストから HTML タグを削除"""
        # <div>, <br>, <span> などのタグと属性を削除
        text = re.sub(r"<[^>]+>", "", text)
        # 連続した空白を1つに統一
        text = re.sub(r"\s+", " ", text).strip()
        # HTMLエンティティをデコード
        text = text.replace("&nbsp;", " ")
        return text
