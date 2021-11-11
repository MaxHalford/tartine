import enum
import re
import string
from collections import UserDict, UserList
from dataclasses import dataclass
from typing import Dict, List, Optional, Tuple


__all__ = ["spread"]


class _Flavor(enum.Enum):
    PYGSHEETS = "pygsheets"


def _column_letter(n: int) -> str:
    """

    >>> _column_letter(0)
    'A'

    >>> _column_letter(1)
    'B'

    >>> _column_letter(25)
    'Z'

    >>> _column_letter(26)
    'AA'

    >>> _column_letter(27)
    'AB'

    """
    code = ""
    n += 1
    while n > 0:
        n, mod = divmod(n - 1, 26)
        code = chr(65 + mod) + code
    return code


def _is_variable(expr: str) -> bool:
    """

    >>> _is_variable("foo")
    False

    >>> _is_variable("@foo")
    True

    """
    return expr.startswith("@")


def _is_formula(expr: str) -> bool:
    """

    >>> _is_formula("foo")
    False

    >>> _is_formula("@foo")
    False

    >>> _is_formula("= @foo + @bar")
    True

    """
    return expr.startswith("=")


def _is_named_formula(expr):
    """

    >>> _is_named_formula("foo")
    False

    >>> _is_named_formula("@foo")
    False

    >>> _is_named_formula("= @foo + @bar")
    False

    >>> _is_named_formula("bar = 42 * @foo")
    True

    """
    return bool(re.match(r"[\w_\.]+ = ", expr))


def normalize_expression(expr: str) -> str:
    """

    >>> normalize_expression(" 1")
    '1'

    >>> normalize_expression("foo  = 2 * 3 ")
    'foo = 2 * 3'

    """
    return re.sub(r"\s+", " ", expr).strip()


@dataclass
class _Cell:
    r: int
    c: int
    expr: str

    @property
    def address(self) -> str:
        """

        >>> _Cell(0, 0, '').address
        'A1'

        >>> _Cell(0, 5, '').address
        'F1'

        >>> _Cell(1000, 500, '').address
        'SG1001'

        """
        return f"{_column_letter(self.c)}{self.r + 1}"

    def as_pygsheets(self, data, annotate: bool) -> "pygsheets.Cell":
        import pygsheets

        expr = _bake_expression(self.expr, data)
        cell = pygsheets.Cell(pos=self.address, val=expr)

        if annotate and _is_named_formula(self.expr):
            name = self.expr.split(" = ")[0]
            cell.note = name

        if _is_formula(expr):
            cell.formula = expr

        return cell


def _bake_expression(expr: str, data: dict) -> str:

    # Replace variables
    for pattern in [r"@'(?P<name>.+)'", r"@(?P<name>\w+)"]:
        expr = re.sub(pattern, lambda m: data[m.group("name")], expr)

    # Remove names from named variables
    for pattern in [r"'.+' = ", r"\w+ = "]:
        expr = re.sub(pattern, "= ", expr)

    return expr


def spread(
    template: List,
    data: Optional[Dict],
    start_row: int = 0,
    flavor: str = _Flavor.PYGSHEETS,
    annotate: bool = False,
) -> Tuple[List[_Cell], int]:
    """Spread data into cells.

    Parameters
    ----------
    template
        A list of expressions which determines how the cells are layed out.
    data.
        A dictionary of data to render.
    start_row
        The row number where the layout begins. Zero-based.
    flavor
        Determines what kinds of cells to generate. Only the `pygsheets` flavor is supported right now.
    annotate
        Determines whether or not to attach notes to cells which contain named formulas.

    Returns
    -------
    The list of cells.
    The number of rows which the cells span over.

    """

    data = data or {}

    table = [
        [
            _Cell(r + start_row, c, normalize_expression(expr))
            for r, expr in enumerate([col] if isinstance(col, str) else col)
        ]
        for c, col in enumerate(template)
    ]

    # We're going to add the positions of the named variables to the data
    data = data.copy()
    for c, col in enumerate(table):
        for r, cell in enumerate(col):
            if _is_named_formula(cell.expr):
                name = cell.expr.split(" = ")[0]
                data[name] = cell.address

    if flavor == _Flavor.PYGSHEETS.value:
        cells = [
            cell.as_pygsheets(data=data, annotate=annotate)
            for col in table
            for cell in col
        ]
    else:
        raise ValueError(
            f"Unknown flavor. Available options: {', '.join(f.value for f in _Flavor)}"
        )

    n_rows = max(map(len, table))
    return cells, n_rows
