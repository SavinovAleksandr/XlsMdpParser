"""Tokenizer for MDP formula expressions."""
from __future__ import annotations

import re
from dataclasses import dataclass
from enum import Enum, auto

from ..normalization import strip_formula_prefix


class TokenKind(Enum):
    NUMBER = auto()
    IDENT = auto()
    PLUS = auto()
    MINUS = auto()
    STAR = auto()
    SLASH = auto()
    CARET = auto()
    LPAREN = auto()
    RPAREN = auto()
    COMMA = auto()
    EQ = auto()
    NE = auto()
    LT = auto()
    LE = auto()
    GT = auto()
    GE = auto()
    EOF = auto()


@dataclass
class Token:
    kind: TokenKind
    value: str
    pos: int


IDENT_RE = re.compile(r"[A-Za-zА-Яа-яЁё_][A-Za-zА-Яа-яЁё0-9_]*")
NUMBER_RE = re.compile(r"\d+(?:\.\d+)?")


def _matching_paren(source: str, open_pos: int) -> int:
    depth = 0
    for pos in range(open_pos, len(source)):
        if source[pos] == "(":
            depth += 1
        elif source[pos] == ")":
            depth -= 1
            if depth == 0:
                return pos
    raise ValueError("Expected ')' after CASE")


def _split_arguments(source: str) -> list[str]:
    parts: list[str] = []
    depth = 0
    start = 0
    for pos, char in enumerate(source):
        if char == "(":
            depth += 1
        elif char == ")":
            depth -= 1
        elif char == "," and depth == 0:
            parts.append(source[start:pos].strip())
            start = pos + 1
    parts.append(source[start:].strip())
    return parts


def _call_arguments(source: str, name: str) -> list[str] | None:
    match = re.fullmatch(rf"\s*{name}\s*\((.*)\)\s*", source, re.IGNORECASE | re.DOTALL)
    return _split_arguments(match.group(1)) if match else None


def _expand_case_call(content: str) -> str:
    args = _split_arguments(content)
    if len(args) < 3:
        raise ValueError("CASE requires a selector and at least one IS/RETURN pair")
    selector = args[0]
    branches: list[tuple[list[str], str]] = []
    default = "9999"
    pos = 1
    while pos < len(args):
        values = _call_arguments(args[pos], "is")
        if values is None:
            fallback = _call_arguments(args[pos], "return")
            if fallback and len(fallback) == 1:
                default = fallback[0]
                pos += 1
                continue
            raise ValueError(f"Expected IS(...) in CASE, got {args[pos]!r}")
        if pos + 1 >= len(args):
            raise ValueError("Expected RETURN(...) after IS(...)")
        returned = _call_arguments(args[pos + 1], "return")
        if not returned or len(returned) != 1:
            raise ValueError("Expected one expression inside RETURN(...)")
        branches.append((values, returned[0]))
        pos += 2

    result = default
    for values, returned in reversed(branches):
        comparisons = [f"({selector})==({value})" for value in values]
        condition = " or ".join(comparisons)
        result = f"if(({condition}),({returned}),({result}))"
    return result


def expand_case_expressions(source: str) -> str:
    """Convert CASE(x, IS(...), RETURN(...)) constructs to nested IF calls."""
    while True:
        match = re.search(r"(?<![A-Za-zА-Яа-яЁё0-9_])case\s*\(", source, re.IGNORECASE)
        if not match:
            return source
        open_pos = source.find("(", match.start())
        close_pos = _matching_paren(source, open_pos)
        replacement = _expand_case_call(source[open_pos + 1 : close_pos])
        source = source[: match.start()] + replacement + source[close_pos + 1 :]


def normalize_formula_source(source: str) -> str:
    """Normalize identifiers exported from human-readable Excel factor names."""
    source = strip_formula_prefix(source)
    source = expand_case_expressions(source)
    # The workbook generator uses a single ampersand/vertical bar for boolean
    # operations, while our expression grammar uses AND/OR.
    source = re.sub(r"\s*&+\s*", " and ", source)
    source = re.sub(r"\s*\|+\s*", " or ", source)
    # Symbols that occur inside equipment names are converted to identifier-safe
    # text. Parentheses, brackets, commas and other punctuation are normally
    # already replaced by underscores by the workbook generator.
    source = source.replace("№", "N")

    # Some formulas retain spaces from the information-sheet label, while other
    # workbooks replace them with underscores. Protect logical words, then join
    # adjacent identifier words so both representations produce the same AST.
    protected: dict[str, str] = {}

    def protect_logic(match: re.Match[str]) -> str:
        marker = f"§{len(protected)}§"
        protected[marker] = match.group(0)
        return f" {marker} "

    source = re.sub(
        r"(?<![A-Za-zА-Яа-яЁё0-9_])(and|or|not|и|или|не)(?![A-Za-zА-Яа-яЁё0-9_])",
        protect_logic,
        source,
        flags=re.IGNORECASE,
    )
    source = re.sub(
        r"(?<=[A-Za-zА-Яа-яЁё0-9_])\s+(?=[A-Za-zА-Яа-яЁё_])",
        "_",
        source,
    )
    for marker, value in protected.items():
        source = source.replace(marker, value)
    return source


class Tokenizer:
    def __init__(self, source: str) -> None:
        source = normalize_formula_source(source)
        # Excel uses either semicolons + decimal commas (Russian locale), or
        # commas + decimal points.  Treating every comma between digits as a
        # decimal separator turned IF(x,313,303) into a two-argument call.
        if ";" in source:
            source = re.sub(r"(?<=\d),(?=\d)", ".", source).replace(";", ",")
        # Typographic dashes occur inside long parameter identifiers.  The
        # workbook convention already uses underscores as spaces, so an
        # underscore is the least surprising normalized representation.
        source = source.replace("–", "_").replace("—", "_").replace("−", "-")
        self.source = source
        self.pos = 0
        self.tokens: list[Token] = []
        self._tokenize()

    def _peek(self, offset: int = 0) -> str:
        idx = self.pos + offset
        return self.source[idx] if idx < len(self.source) else ""

    def _advance(self, count: int = 1) -> None:
        self.pos += count

    def _skip_ws(self) -> None:
        while self.pos < len(self.source) and self.source[self.pos].isspace():
            self.pos += 1

    def _tokenize(self) -> None:
        while True:
            self._skip_ws()
            start = self.pos
            if self.pos >= len(self.source):
                self.tokens.append(Token(TokenKind.EOF, "", start))
                break
            ch = self.source[self.pos]
            two = self.source[self.pos : self.pos + 2]
            if two in ("<=", ">=", "<>", "!=", "=="):
                kind = {
                    "<=": TokenKind.LE,
                    ">=": TokenKind.GE,
                    "<>": TokenKind.NE,
                    "!=": TokenKind.NE,
                    "==": TokenKind.EQ,
                }[two]
                self.tokens.append(Token(kind, two, start))
                self._advance(2)
                continue
            if ch in "+":
                self.tokens.append(Token(TokenKind.PLUS, ch, start))
                self._advance()
                continue
            if ch in "-":
                self.tokens.append(Token(TokenKind.MINUS, ch, start))
                self._advance()
                continue
            if ch in "*·":
                self.tokens.append(Token(TokenKind.STAR, ch, start))
                self._advance()
                continue
            if ch in "/":
                self.tokens.append(Token(TokenKind.SLASH, ch, start))
                self._advance()
                continue
            if ch in "^":
                self.tokens.append(Token(TokenKind.CARET, ch, start))
                self._advance()
                continue
            if ch in "(":
                self.tokens.append(Token(TokenKind.LPAREN, ch, start))
                self._advance()
                continue
            if ch in ")":
                self.tokens.append(Token(TokenKind.RPAREN, ch, start))
                self._advance()
                continue
            if ch in ",":
                self.tokens.append(Token(TokenKind.COMMA, ch, start))
                self._advance()
                continue
            if ch in "<":
                self.tokens.append(Token(TokenKind.LT, ch, start))
                self._advance()
                continue
            if ch in ">":
                self.tokens.append(Token(TokenKind.GT, ch, start))
                self._advance()
                continue
            if ch in "=":
                self.tokens.append(Token(TokenKind.EQ, ch, start))
                self._advance()
                continue
            num_match = NUMBER_RE.match(self.source, self.pos)
            if num_match:
                self.tokens.append(Token(TokenKind.NUMBER, num_match.group(0), start))
                self._advance(len(num_match.group(0)))
                continue
            ident_match = IDENT_RE.match(self.source, self.pos)
            if ident_match:
                self.tokens.append(Token(TokenKind.IDENT, ident_match.group(0), start))
                self._advance(len(ident_match.group(0)))
                continue
            raise ValueError(f"Unexpected character {ch!r} at position {self.pos}")

    def __iter__(self):
        return iter(self.tokens)
