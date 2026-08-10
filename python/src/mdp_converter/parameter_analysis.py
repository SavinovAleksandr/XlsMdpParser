"""Mode parameter and factor extraction from formulas."""
from __future__ import annotations

from .expression.ast import AstNode, CompareNode, IfNode, NumberNode, VariableNode
from .expression.evaluator import collect_variables, variables_in_conditions, variables_in_expressions
from .models import Factor, ModeOption, ModeParam
from .normalization import factor_key, merge_factor_names, norm


def analyze_parameters(
    ast_nodes: list[AstNode | None],
    info_factors: list[str],
) -> tuple[list[ModeParam], list[Factor]]:
    condition_vars: set[str] = set()
    expression_vars: set[str] = set()
    boolean_vars: set[str] = set()
    comparison_values: dict[str, set[float]] = {}

    for ast in ast_nodes:
        if ast is None:
            continue
        condition_vars |= variables_in_conditions(ast)
        expression_vars |= variables_in_expressions(ast)
        _scan_boolean(ast, boolean_vars)
        _scan_comparisons(ast, comparison_values)

    mode_names = sorted(condition_vars)
    expression_factor_names = sorted(expression_vars - condition_vars)
    # Some workbooks describe a factor generically on the information sheet
    # (for example ``Pнб``), while formulas use its full direction-qualified
    # name (``Pнб__Коноша_–_Вельск__Р``).  Keeping both creates a dead input
    # that is never referenced by a formula.  Prefer the precise formula name.
    precise_keys = [factor_key(name) for name in expression_factor_names]
    useful_info_factors = [
        name
        for name in info_factors
        if factor_key(name) in precise_keys
    ]
    factor_names = merge_factor_names(useful_info_factors + expression_factor_names)

    mode_params: list[ModeParam] = []
    for name in mode_names:
        if name in boolean_vars or _looks_boolean(name, comparison_values.get(name, set())):
            mode_params.append(ModeParam(name=name, kind="bool", default="0"))
            continue
        values = comparison_values.get(name, set())
        if values:
            opts = _build_select_options(name, values)
            default = (
                "1"
                if "сезон" in name.lower()
                and "аопо" in name.lower()
                and any(option.value == "1" for option in opts)
                else (opts[0].value if opts else "0")
            )
            mode_params.append(
                ModeParam(
                    name=name,
                    kind="select",
                    options=opts,
                    default=default,
                )
            )
        else:
            mode_params.append(ModeParam(name=name, kind="number", default="0"))

    factors = [Factor(name=n) for n in factor_names if factor_key(n) not in {factor_key(m) for m in mode_names}]
    return mode_params, factors


def _looks_boolean(name: str, values: set[float]) -> bool:
    lower = name.lower()
    if "сезон" in lower and "аопо" in lower:
        return False
    if lower.startswith("фиксация") or lower.startswith("fix"):
        return True
    return values <= {0.0, 1.0} and bool(values)


def _scan_boolean(node: AstNode, out: set[str]) -> None:
    if isinstance(node, IfNode):
        if isinstance(node.cond, VariableNode):
            out.add(node.cond.name)
    for child in _iter_children(node):
        _scan_boolean(child, out)


def _scan_comparisons(node: AstNode, out: dict[str, set[float]]) -> None:
    if isinstance(node, CompareNode):
        if isinstance(node.left, VariableNode) and isinstance(node.right, NumberNode):
            _add_boundary_values(out, node.left.name, node.op, node.right.value)
        elif isinstance(node.right, VariableNode) and isinstance(node.left, NumberNode):
            reverse = {"<": ">", "<=": ">=", ">": "<", ">=": "<=", "==": "==", "<>": "<>"}
            _add_boundary_values(out, node.right.name, reverse.get(node.op, node.op), node.left.value)
    for child in _iter_children(node):
        _scan_comparisons(child, out)


def _add_boundary_values(out: dict[str, set[float]], name: str, op: str, value: float) -> None:
    """Add representative values for both sides of a comparison boundary."""
    values = out.setdefault(name, set())
    values.add(value)
    if op in (">", ">="):
        values.add(value - 1 if op == ">=" else value + 1)
    elif op in ("<", "<="):
        values.add(value - 1 if op == "<" else value + 1)


def _build_select_options(name: str, values: set[float]) -> list[ModeOption]:
    lower = name.lower()
    if "сезон" in lower and "аопо" in lower:
        sorted_vals = sorted(values)
    else:
        sorted_vals = sorted(values, reverse=True)
    options: list[ModeOption] = []
    for val in sorted_vals:
        iv = int(val) if val == int(val) else val
        label = _label_for_value(name, iv)
        options.append(ModeOption(value=str(iv), label=label))
    if not options:
        options.append(ModeOption(value="0", label="0"))
    return options


def _label_for_value(name: str, value: int | float) -> str:
    lower = name.lower()
    if "сезон" in lower and "аопо" in lower:
        return f"Группа {int(value)}"
    if "ртд" in lower or "reactor" in lower:
        return "Включен" if value else "Отключен"
    if "кол" in lower or "ген" in lower or "блок" in lower or "тг" in lower:
        n = int(value)
        suffix = "генератор" if n == 1 else "генератора" if 2 <= n <= 4 else "генераторов"
        return f"{n} {suffix}"
    return str(value)


def _iter_children(node: AstNode):
    from .expression.ast import BinaryNode, FunctionNode, LogicNode, UnaryNode

    if isinstance(node, (UnaryNode,)):
        yield node.arg
    elif isinstance(node, (BinaryNode, CompareNode)):
        yield node.left
        yield node.right
    elif isinstance(node, LogicNode):
        yield from node.args
    elif isinstance(node, FunctionNode):
        yield from node.args
    elif isinstance(node, IfNode):
        yield node.cond
        yield node.then_branch
        yield node.else_branch
