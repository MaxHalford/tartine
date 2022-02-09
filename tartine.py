import enum
import re
from dataclasses import dataclass
from typing import Callable, Dict, List, Optional, Tuple, Union


__all__ = ["spread", "spread_dataframe", "unspread_dataframe"]


class Flavor(enum.Enum):
    PYGSHEETS = "pygsheets"


Template = Dict[str, Union[str, Tuple[str, ...]]]


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


def _is_named_formula(expr: str) -> bool:
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


def _normalize_expression(expr: str) -> str:
    """

    >>> _normalize_expression(" 1")
    '1'

    >>> _normalize_expression("foo  = 2 * 3 ")
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

    def as_pygsheets(self, data) -> "pygsheets.Cell":
        import pygsheets

        expr = _bake_expression(self.expr, data)
        cell = pygsheets.Cell(pos=self.address, val=expr)

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
    template: Template,
    data: Optional[Dict],
    flavor: Flavor,
    postprocess: Optional[Callable] = None,
    start_at: int = 0,
) -> Tuple[List[Union["pygsheets.Cell"]], int]:
    """Spread data into cells.

    Parameters
    ----------
    template
        A list of expressions which determines how the cells are layed out.
    data
        A dictionary of data to render.
    flavor
        Determines what kind of cells to generate.
    postprocess
        An optional function to call for each cell once it has been created.
    start_at
        The row number where the layout begins. Zero-based.

    Returns
    -------
    cells
        The list of cells.
    n_rows
        The number of rows which the cells span over.

    """

    data = data or {}

    table = [
        [
            _Cell(r + start_at, c, _normalize_expression(expr))
            for r, expr in enumerate([col] if isinstance(col, str) else col)
        ]
        for c, col in enumerate(template)
    ]

    # We're going to add the positions of the named variables to the data
    data = data.copy()
    cell_names = {}
    for c, col in enumerate(table):
        for r, cell in enumerate(col):
            if _is_named_formula(cell.expr):
                name = cell.expr.split(" = ")[0]
                data[name] = cell.address
                cell_names[len(cell_names)] = name
            elif _is_variable(cell.expr):
                cell_names[len(cell_names)] = cell.expr[1:]
            else:
                cell_names[len(cell_names)] = None

    if flavor == Flavor.PYGSHEETS.value:
        cells = [cell.as_pygsheets(data=data) for col in table for cell in col]
    else:
        raise ValueError(
            f"Unknown flavor {flavor}. Available options: {', '.join(f.value for f in Flavor)}"
        )

    if postprocess:
        for i, cell in enumerate(cells):
            cells[i] = postprocess(cell, cell_names[i])

    n_rows = max(map(len, table))

    return cells, n_rows


def spread_dataframe(
    template: Template,
    df: "pd.DataFrame",
    flavor: Flavor,
    postprocess: Optional[Callable] = None,
) -> List[Union["pygsheets.Cell"]]:
    """Spread a dataframe into cells.

    Parameters
    ----------
    template
        A list of expressions which determines how the cells are layed out.
    df
        A dataframe to render.
    flavor
        Determines what kind of cells to generate.
    postprocess
        An optional function to call for each cell once it has been created.

    Returns
    -------
    cells
        The list of cells.

    """

    cells, nrows = spread(template.keys(), data=None, flavor=flavor)

    for card_set in df.to_dict("records"):
        _cells, _nrows = spread(
            template=template.values(),
            data=card_set,
            start_at=nrows,
            flavor=flavor,
            postprocess=postprocess,
        )

        cells += _cells
        nrows += _nrows

    return cells


def unspread_dataframe(template: Template, df: "pd.DataFrame") -> "pd.DataFrame":
    """Unspread a dataframe into a flat dataframe.

    Parameters
    ----------
    template
        A list of expressions which determines how the cells are layed out.
    df
        A dataframe to unspread. Typically this may be the output of the `spread` function once
        it has been dumped into a sheet.

    Returns
    -------
    flat_df
        The flattened dataframe.

    """

    import pandas as pd

    n_rows_in_template = max(
        1 if isinstance(col, str) else len(col) for col in template.values()
    )

    flat_rows = []

    for k in range(n_rows_in_template, len(df), n_rows_in_template):

        group = df[k - n_rows_in_template : k]

        flat_rows.append(
            {
                var: group[col].iloc[i]
                for col, variables in template.items()
                for i, var in enumerate(variables)
                if var
            }
        )

    return pd.DataFrame(flat_rows)
