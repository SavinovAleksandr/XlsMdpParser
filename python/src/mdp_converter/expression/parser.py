"""Recursive descent parser for MDP formulas."""
from __future__ import annotations

from .ast import (
    AstNode,
    BinaryNode,
    CompareNode,
    FunctionNode,
    IfNode,
    LogicNode,
    NumberNode,
    UnaryNode,
    VariableNode,
)
from .tokenizer import Token, TokenKind, Tokenizer


class Parser:
    def __init__(self, source: str) -> None:
        self.tokens = list(Tokenizer(source))
        self.pos = 0

    def parse(self) -> AstNode:
        node = self._parse_or()
        if self._current().kind != TokenKind.EOF:
            raise ValueError(f"Unexpected token {self._current().value!r}")
        return node

    def _current(self) -> Token:
        return self.tokens[self.pos]

    def _advance(self) -> Token:
        tok = self.tokens[self.pos]
        self.pos += 1
        return tok

    def _match(self, kind: TokenKind) -> bool:
        if self._current().kind == kind:
            self._advance()
            return True
        return False

    def _parse_or(self) -> AstNode:
        left = self._parse_and()
        while self._is_logic("or", "или"):
            self._advance()
            right = self._parse_and()
            if isinstance(left, LogicNode) and left.op == "or":
                left.args.append(right)
            else:
                left = LogicNode("or", [left, right])
        return left

    def _parse_and(self) -> AstNode:
        left = self._parse_not()
        while self._is_logic("and", "и"):
            self._advance()
            right = self._parse_not()
            if isinstance(left, LogicNode) and left.op == "and":
                left.args.append(right)
            else:
                left = LogicNode("and", [left, right])
        return left

    def _is_logic(self, *names: str) -> bool:
        tok = self._current()
        return tok.kind == TokenKind.IDENT and tok.value.lower() in names

    def _parse_not(self) -> AstNode:
        if self._is_logic("not", "не"):
            self._advance()
            return UnaryNode("not", self._parse_not())
        return self._parse_comparison()

    def _parse_comparison(self) -> AstNode:
        left = self._parse_additive()
        ops = {
            TokenKind.EQ: "==",
            TokenKind.NE: "<>",
            TokenKind.LT: "<",
            TokenKind.LE: "<=",
            TokenKind.GT: ">",
            TokenKind.GE: ">=",
        }
        if self._current().kind in ops:
            op = ops[self._current().kind]
            self._advance()
            right = self._parse_additive()
            return CompareNode(op, left, right)
        return left

    def _parse_additive(self) -> AstNode:
        left = self._parse_multiplicative()
        while self._current().kind in (TokenKind.PLUS, TokenKind.MINUS):
            op = self._advance().value
            right = self._parse_multiplicative()
            left = BinaryNode(op, left, right)
        return left

    def _parse_multiplicative(self) -> AstNode:
        left = self._parse_power()
        while self._current().kind in (TokenKind.STAR, TokenKind.SLASH):
            op = "*" if self._current().kind == TokenKind.STAR else "/"
            self._advance()
            right = self._parse_power()
            left = BinaryNode(op, left, right)
        return left

    def _parse_power(self) -> AstNode:
        left = self._parse_unary()
        if self._current().kind == TokenKind.CARET:
            self._advance()
            right = self._parse_power()
            return BinaryNode("^", left, right)
        return left

    def _parse_unary(self) -> AstNode:
        if self._current().kind == TokenKind.MINUS:
            self._advance()
            return UnaryNode("-", self._parse_unary())
        return self._parse_primary()

    def _parse_primary(self) -> AstNode:
        tok = self._current()
        if tok.kind == TokenKind.NUMBER:
            self._advance()
            value = float(tok.value.replace(",", "."))
            return NumberNode(value)
        if tok.kind == TokenKind.IDENT:
            name = tok.value
            lower = name.lower()
            self._advance()
            if self._current().kind == TokenKind.LPAREN:
                return self._parse_call(lower)
            return VariableNode(name)
        if tok.kind == TokenKind.LPAREN:
            self._advance()
            node = self._parse_or()
            if not self._match(TokenKind.RPAREN):
                raise ValueError("Expected ')'")
            return node
        raise ValueError(f"Unexpected token {tok.value!r}")

    def _parse_call(self, name: str) -> AstNode:
        self._advance()  # (
        args: list[AstNode] = []
        if self._current().kind != TokenKind.RPAREN:
            args.append(self._parse_or())
            while self._match(TokenKind.COMMA):
                args.append(self._parse_or())
        if not self._match(TokenKind.RPAREN):
            raise ValueError("Expected ')' after function arguments")
        if name == "if" and len(args) == 3:
            return IfNode(args[0], args[1], args[2])
        if name in ("abs", "min", "max"):
            return FunctionNode(name, args)
        raise ValueError(f"Unknown function {name!r}")


def parse_expression(source: str) -> AstNode:
    if not source or not source.strip():
        raise ValueError("Empty expression")
    return Parser(source).parse()
