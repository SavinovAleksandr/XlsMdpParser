"""AST evaluation and analysis."""
from __future__ import annotations

import math
from typing import Any

from .ast import (
    AstNode,
    BinaryNode,
    CompareNode,
    FunctionNode,
    IfNode,
    LogicNode,
    NumberNode,
    RESERVED,
    UnaryNode,
    VariableNode,
)


def _num(value: Any) -> float:
    if isinstance(value, bool):
        return 1.0 if value else 0.0
    try:
        return float(value)
    except (TypeError, ValueError):
        return float("nan")


def _lookup(name: str, env: dict[str, Any]) -> Any:
    if name in env:
        return env[name]
    for key, value in env.items():
        if key.replace("_", " ") == name.replace("_", " "):
            return value
        if key.replace(" ", "_") == name.replace(" ", "_"):
            return value
    return 0


def evaluate(node: AstNode | dict, env: dict[str, Any]) -> float:
    if isinstance(node, dict):
        return evaluate(_from_dict(node), env)
    if isinstance(node, NumberNode):
        return node.value
    if isinstance(node, VariableNode):
        return _num(_lookup(node.name, env))
    if isinstance(node, UnaryNode):
        val = evaluate(node.arg, env)
        if node.op == "-":
            return -val
        if node.op == "not":
            return 0.0 if val else 1.0
    if isinstance(node, BinaryNode):
        left = evaluate(node.left, env)
        right = evaluate(node.right, env)
        if node.op == "+":
            return left + right
        if node.op == "-":
            return left - right
        if node.op == "*":
            return left * right
        if node.op == "/":
            return left / right if right else float("nan")
        if node.op == "^":
            return left ** right
    if isinstance(node, CompareNode):
        left = evaluate(node.left, env)
        right = evaluate(node.right, env)
        if node.op == "==":
            return 1.0 if left == right else 0.0
        if node.op == "<>":
            return 1.0 if left != right else 0.0
        if node.op == "<":
            return 1.0 if left < right else 0.0
        if node.op == "<=":
            return 1.0 if left <= right else 0.0
        if node.op == ">":
            return 1.0 if left > right else 0.0
        if node.op == ">=":
            return 1.0 if left >= right else 0.0
    if isinstance(node, LogicNode):
        if node.op == "and":
            return 1.0 if all(evaluate(arg, env) for arg in node.args) else 0.0
        if node.op == "or":
            return 1.0 if any(evaluate(arg, env) for arg in node.args) else 0.0
    if isinstance(node, FunctionNode):
        args = [evaluate(arg, env) for arg in node.args]
        if node.name == "abs":
            return abs(args[0]) if args else float("nan")
        if node.name == "min":
            return min(args) if args else float("nan")
        if node.name == "max":
            return max(args) if args else float("nan")
    if isinstance(node, IfNode):
        return evaluate(node.then_branch if evaluate(node.cond, env) else node.else_branch, env)
    return float("nan")


def active_branch(node: AstNode | dict, env: dict[str, Any]) -> AstNode:
    """Resolve conditional nodes anywhere inside the selected formula."""
    if isinstance(node, dict):
        node = _from_dict(node)
    if isinstance(node, IfNode):
        if evaluate(node.cond, env):
            return active_branch(node.then_branch, env)
        return active_branch(node.else_branch, env)
    if isinstance(node, UnaryNode):
        return UnaryNode(node.op, active_branch(node.arg, env))
    if isinstance(node, BinaryNode):
        return BinaryNode(node.op, active_branch(node.left, env), active_branch(node.right, env))
    if isinstance(node, CompareNode):
        return CompareNode(node.op, active_branch(node.left, env), active_branch(node.right, env))
    if isinstance(node, LogicNode):
        return LogicNode(node.op, [active_branch(arg, env) for arg in node.args])
    if isinstance(node, FunctionNode):
        return FunctionNode(node.name, [active_branch(arg, env) for arg in node.args])
    return node


def collect_variables(node: AstNode | dict | None) -> set[str]:
    if node is None:
        return set()
    if isinstance(node, dict):
        node = _from_dict(node)
    if isinstance(node, VariableNode):
        return {node.name}
    names: set[str] = set()
    for child in _children(node):
        names |= collect_variables(child)
    return names


def variables_in_conditions(node: AstNode | dict | None) -> set[str]:
    if node is None:
        return set()
    if isinstance(node, dict):
        node = _from_dict(node)
    names: set[str] = set()
    if isinstance(node, IfNode):
        names |= collect_variables(node.cond)
        names |= variables_in_conditions(node.then_branch)
        names |= variables_in_conditions(node.else_branch)
    elif isinstance(node, (CompareNode, LogicNode)):
        names |= collect_variables(node)
    else:
        for child in _children(node):
            names |= variables_in_conditions(child)
    return names


def variables_in_expressions(node: AstNode | dict | None) -> set[str]:
    if node is None:
        return set()
    if isinstance(node, dict):
        node = _from_dict(node)
    if isinstance(node, IfNode):
        return variables_in_expressions(node.then_branch) | variables_in_expressions(node.else_branch)
    if isinstance(node, (CompareNode, LogicNode)):
        return set()
    if isinstance(node, VariableNode):
        return {node.name}
    names: set[str] = set()
    for child in _children(node):
        names |= variables_in_expressions(child)
    return names


def format_expression(node: AstNode | dict) -> str:
    if isinstance(node, dict):
        node = _from_dict(node)
    if isinstance(node, NumberNode):
        if node.value == int(node.value):
            return str(int(node.value))
        return str(round(node.value, 6)).rstrip("0").rstrip(".")
    if isinstance(node, VariableNode):
        return node.name
    if isinstance(node, UnaryNode):
        if node.op == "not":
            return f"NOT {format_expression(node.arg)}"
        return f"-{format_expression(node.arg)}"
    if isinstance(node, BinaryNode):
        left = format_expression(node.left)
        right = format_expression(node.right)
        sym = "×" if node.op == "*" else "÷" if node.op == "/" else node.op
        return f"{left} {sym} {right}"
    if isinstance(node, CompareNode):
        return f"{format_expression(node.left)} {node.op} {format_expression(node.right)}"
    if isinstance(node, LogicNode):
        join = " AND " if node.op == "and" else " OR "
        return join.join(format_expression(a) for a in node.args)
    if isinstance(node, FunctionNode):
        args = ", ".join(format_expression(a) for a in node.args)
        return f"{node.name.upper()}({args})"
    if isinstance(node, IfNode):
        return f"IF({format_expression(node.cond)}, {format_expression(node.then_branch)}, {format_expression(node.else_branch)})"
    return ""


def is_valid_result(value: float, sentinel: float = 999.0) -> bool:
    if not math.isfinite(value):
        return False
    if abs(value - sentinel) < 1e-9:
        return False
    return True


def minimum_criterion(values: list[tuple[int, float]]) -> int | None:
    valid = [(num, val) for num, val in values if is_valid_result(val)]
    if not valid:
        return None
    return min(valid, key=lambda item: item[1])[0]


def _children(node: AstNode) -> list[AstNode]:
    if isinstance(node, (UnaryNode,)):
        return [node.arg]
    if isinstance(node, (BinaryNode, CompareNode)):
        return [node.left, node.right]
    if isinstance(node, LogicNode):
        return node.args
    if isinstance(node, FunctionNode):
        return node.args
    if isinstance(node, IfNode):
        return [node.cond, node.then_branch, node.else_branch]
    return []


def _from_dict(data: dict) -> AstNode:
    kind = data["type"]
    if kind == "num":
        return NumberNode(float(data["value"]))
    if kind == "var":
        return VariableNode(data["name"])
    if kind == "un":
        return UnaryNode(data["op"], _from_dict(data["arg"]))
    if kind == "bin":
        return BinaryNode(data["op"], _from_dict(data["left"]), _from_dict(data["right"]))
    if kind == "cmp":
        return CompareNode(data["op"], _from_dict(data["left"]), _from_dict(data["right"]))
    if kind == "logic":
        return LogicNode(data["op"], [_from_dict(a) for a in data["args"]])
    if kind == "func":
        return FunctionNode(data["name"], [_from_dict(a) for a in data["args"]])
    if kind == "if":
        return IfNode(_from_dict(data["cond"]), _from_dict(data["then"]), _from_dict(data["else"]))
    raise ValueError(f"Unknown AST node {kind!r}")
