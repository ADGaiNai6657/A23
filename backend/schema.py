from collections.abc import Iterator
from typing import Literal

from pydantic import BaseModel, Field


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
    value: str = Field(min_length=1)
    unit: str = ""
    snippet: str = Field(min_length=1)
    scope: str = ""


class RawFacts(BaseModel):
    facts: list[RawFact]


class FillRequest(BaseModel):
    doc: str
    template: str


Mode = Literal["general", "list"]


class ExtractResult(BaseModel):
    doc: str
    mode: Mode
    count: int
    kept: int
