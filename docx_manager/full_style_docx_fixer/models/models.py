from dataclasses import dataclass, field
from typing import List, Dict, Any, Optional
from enum import Enum


class ContentType(Enum):
    SECTION = "section"
    TOC = "toc"
    HEADING1 = "heading1"
    HEADING2 = "heading2"
    HEADING3 = "heading3"
    BODY = "body"
    TABLE = "table"
    FORMULA = "formula"
    IMAGE = "image"


class SectionType(Enum):
    ABSTRACT = "abstract"
    ABSTRACT_EN = "abstract_en"
    CONCLUSION = "conclusion"
    ACKNOWLEDGEMENT = "acknowledgement"
    CUSTOM = "custom"


@dataclass
class ContentItem:
    type: str
    data: Dict[str, Any]


@dataclass
class SectionContent(ContentItem):
    section_type: str
    toc_exclude: bool
    value: str
    title: Optional[str] = None


@dataclass
class TocContent(ContentItem):
    title: str
    toc_title_exclude: bool


@dataclass
class HeadingContent(ContentItem):
    value: str
    toc_exclude: Optional[bool] = None


@dataclass
class BodyContent(ContentItem):
    value: str


@dataclass
class TableContent(ContentItem):
    caption: Optional[str] = None
    data: Optional[List[List[str]]] = None


@dataclass
class FormulaContent(ContentItem):
    omml: Optional[str] = None
    latex: Optional[str] = None
    label: Optional[str] = None


@dataclass
class ImageContent(ContentItem):
    base64: Optional[str] = None
    path: Optional[str] = None
    ext: Optional[str] = None
    caption: Optional[str] = None
    width: Optional[float] = None
    align: Optional[str] = None


@dataclass
class TocEntry:
    title: str
    level: int
    page: str


@dataclass
class PageFooterConfig:
    section: str
    style: str
    start: int


@dataclass
class Reference:
    id: int
    text: str


@dataclass
class Citation:
    ref_id: int
    before: str
    after: str


@dataclass
class UserData:
    _doc: Optional[str] = None
    page_footer_config: List[PageFooterConfig] = field(default_factory=list)
    toc_mode: Optional[str] = None
    toc_entries: List[TocEntry] = field(default_factory=list)
    content: List[Dict[str, Any]] = field(default_factory=list)
    references: List[Reference] = field(default_factory=list)
    citations: List[Citation] = field(default_factory=list)

    def to_dict(self) -> Dict[str, Any]:
        result = {}
        if self._doc:
            result["_doc"] = self._doc
        if self.page_footer_config:
            result["page_footer_config"] = [
                {
                    "section": item.section,
                    "style": item.style,
                    "start": item.start
                }
                for item in self.page_footer_config
            ]
        if self.toc_mode:
            result["toc_mode"] = self.toc_mode
        if self.toc_entries:
            result["toc_entries"] = [
                {
                    "title": entry.title,
                    "level": entry.level,
                    "page": entry.page
                }
                for entry in self.toc_entries
            ]
        if self.content:
            result["content"] = self.content
        if self.references:
            result["references"] = [
                {
                    "id": ref.id,
                    "text": ref.text
                }
                for ref in self.references
            ]
        if self.citations:
            result["citations"] = [
                {
                    "ref_id": cit.ref_id,
                    "before": cit.before,
                    "after": cit.after
                }
                for cit in self.citations
            ]
        return result
