"""AST node definitions for formula expressions."""
from __future__ import annotations

from dataclasses import dataclass
from typing import Any


@dataclass
class NumberNode:
    value: float

    def to_dict(self) -> dict[str, Any]:
        return {"type": "num", "value": self.value}


@dataclass
class VariableNode:
    name: str

    def to_dict(self) -> dict[str, Any]:
        return {"type": "var", "name": self.name}


@dataclass
class UnaryNode:
    op: str
    arg: "AstNode"

    def to_dict(self) -> dict[str, Any]:
        return {"type": "un", "op": self.op, "arg": self.arg.to_dict()}


@dataclass
class BinaryNode:
    op: str
    left: "AstNode"
    right: "AstNode"

    def to_dict(self) -> dict[str, Any]:
        return {
            "type": "bin",
            "op": self.op,
            "left": self.left.to_dict(),
            "right": self.right.to_dict(),
        }


@dataclass
class CompareNode:
    op: str
    left: "AstNode"
    right: "AstNode"

    def to_dict(self) -> dict[str, Any]:
        return {
            "type": "cmp",
            "op": self.op,
            "left": self.left.to_dict(),
            "right": self.right.to_dict(),
        }


@dataclass
class LogicNode:
    op: str
    args: list["AstNode"]

    def to_dict(self) -> dict[str, Any]:
        return {"type": "logic", "op": self.op, "args": [a.to_dict() for a in self.args]}


@dataclass
class FunctionNode:
    name: str
    args: list["AstNode"]

    def to_dict(self) -> dict[str, Any]:
        return {"type": "func", "name": self.name.lower(), "args": [a.to_dict() for a in self.args]}


@dataclass
class IfNode:
    cond: "AstNode"
    then_branch: "AstNode"
    else_branch: "AstNode"

    def to_dict(self) -> dict[str, Any]:
        return {
            "type": "if",
            "cond": self.cond.to_dict(),
            "then": self.then_branch.to_dict(),
            "else": self.else_branch.to_dict(),
        }


AstNode = NumberNode | VariableNode | UnaryNode | BinaryNode | CompareNode | LogicNode | FunctionNode | IfNode

RESERVED = frozenset(
    {
        "if",
        "and",
        "or",
        "not",
        "abs",
        "min",
        "max",
        "true",
        "false",
        "и",
        "или",
        "не",
    }
)
