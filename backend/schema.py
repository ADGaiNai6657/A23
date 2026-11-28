from collections.abc import Iterator
from typing import Literal

from pydantic import BaseModel, Field, field_validator


class Block(BaseModel):
    type: str
    text: str
    locator: dict[str, int | str]


class ParsedDoc(BaseModel):
    file: str
    blocks: list[Block]

    def numbered_blocks(self) -> Iterator[tuple[int, Block]]:
        for position, block in enumerate(self.blocks, 1):
            number = int(block.locator["line"]) if block.type == "text_line" else position
            yield number, block


class RawFact(BaseModel):
    block_no: int
    entity: str = ""
    metric: str = Field(min_length=1)
    value: str = Field(min_length=1, coerce_numbers_to_str=True)
    unit: str = ""
    snippet: str = Field(min_length=1)
    scope: str = ""

    @field_validator("entity", "unit", "scope", mode="before")
    @classmethod
    def optional_text(cls, value):
        return "" if value is None else value


class RawFacts(BaseModel):
    facts: list[object]


class FillRequest(BaseModel):
    doc: str
    template: str
    source_sha256: str | None = Field(default=None, pattern=r"^[0-9a-f]{64}$")
    expected_publication_id: int | None = Field(default=None, ge=1)
    target_period: str | None = Field(default=None, pattern=r"^(?:19|20)\d{2}$")


class FactDecisionRequest(BaseModel):
    entity: str = Field(min_length=1)
    metric: str = Field(min_length=1)
    period: str = ""
    scope: str = ""
    preferred_doc: str | None = None


Mode = Literal["general", "list"]


class ClassificationResult(BaseModel):
    schema_version: Literal[1]
    method: Literal["parser_structure_v1"]
    document_format: Literal["txt", "markdown", "docx", "xlsx"]
    content_family: Literal["text", "document", "tabular"]
    block_types: dict[str, int]
    extracted_value_types: dict[str, int]
    semantic_category: str | None = None


class ExtractResult(BaseModel):
    doc: str
    mode: Mode
    count: int
    kept: int
    rejected: dict[str, int] = Field(default_factory=dict)
    classification: ClassificationResult
