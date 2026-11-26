from schema import Block


def parse_text(data: bytes) -> list[Block]:
    text = data.decode("utf-8-sig")
    return [
        Block(type="text_line", text=line.strip(), locator={"line": number})
        for number, line in enumerate(text.splitlines(), 1)
        if line.strip()
    ]
