import enum
import re
from dataclasses import dataclass
from typing import Callable, Dict, List, Optional, Tuple, Union

import glom


__all__ = ["spread", "spread_dataframe", "unspread_dataframe"]


class Flavor(enum.Enum):
    PYGSHEETS = "pygsheets"


StrExpr = str
FuncExpr = Callable[[dict], StrExpr]
Expr = Union[StrExpr, FuncExpr]
Note = Expr
ExprWithNote = Tuple[Expr, Note]
Template = Dict[
    str,
    Union[
        Union[Expr, ExprWithNote],
        List[Union[Expr, ExprWithNote]],
    ],
]


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

    >>> _is_variable("@foo bar")
    True

    >>> _is_variable("@foo bar.foo.bar")
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


def _normalize_expression(expr: Expr) -> Expr:
    """

    >>> _normalize_expression(" 1")
    '1'

    >>> _normalize_expression("foo  = 2 * 3 ")
    'foo = 2 * 3'

    """
    if isinstance(expr, StrExpr):
        return re.sub(r"\s+", " ", expr).strip()
    return expr


@dataclass
class _Cell:
    r: int
    c: int
    expr: str
    note: Optional[str] = None

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

    def as_pygsheets(
        self, data: dict, replace_missing_with: Optional[str] = None
    ) -> "pygsheets.Cell":
        import pygsheets

        expr = _bake_expression(self.expr, data, replace_missing_with)
        cell = pygsheets.Cell(pos=self.address, val=expr)

        if _is_formula(expr):
            cell.formula = expr

        if self.note is not None:
            cell.note = _bake_expression(self.note, data)

        return cell


def _replace_variables(
    str_expr: StrExpr, data: dict, replace_missing_with: Optional[str] = None
):
    """

    >>> _replace_variables('@foo', {'foo': 42})
    '42'

    >>> _replace_variables('2 * @foo', {'foo': 42})
    '2 * 42'

    >>> _replace_variables('2 * @foo bar', {'foo bar': 42})
    '2 * 42'

    >>> _replace_variables('2 * @foo bar - @foo bar', {'foo bar': 42})
    '2 * 42 - 42'

    >>> _replace_variables(
    ...     '2 * @duh - @baz.foo bar',
    ...     {'baz': {'foo bar': 42}},
    ...     replace_missing_with=10
    ... )
    '2 * 10 - 42'

    """

    def replace(match: re.Match) -> str:
        try:
            return str(glom.glom(data, match.group("glom_spec")))
        except KeyError as e:
            if replace_missing_with is not None:
                return str(replace_missing_with)
            raise e

    pattern = r"@(?P<glom_spec>[\w\s\.]*\w)"
    return re.sub(pattern, replace, str_expr)


def _bake_expression(
    expr: Expr, data: dict, replace_missing_with: Optional[str] = None
) -> StrExpr:
    """

    >>> _bake_expression('@a * x + @b', {'a': '3', 'b': '8'})
    '3 * x + 8'

    Values don't have to be strings.

    >>> _bake_expression('@a * x + @b', {'a': 3, 'b': 8})
    '3 * x + 8'

    Nested values can be handled via glom's syntax.

    >>> data = {'coeffs': {'a': 3, 'b': 8}}
    >>> _bake_expression('@coeffs.a * x + @coeffs.b', data)
    '3 * x + 8'

    This also works for formulas.

    >>> data = {
    ...     "rarity": {
    ...         "common": 50,
    ...         "rare cards": 35,
    ...         "epic": 24,
    ...         "legendary": 26
    ...     }
    ... }
    >>> expr = "@rarity.common + @rarity.rare cards + @rarity.epic + @rarity.legendary"
    >>> _bake_expression(expr, data)
    '50 + 35 + 24 + 26'

    You can also use arbitrary functions if string expressions are not enough.

    >>> data = {'a': 3, 'b': 8}
    >>> expr = lambda x: f"{x['a']} * x + {x['b']}"
    >>> _bake_expression(expr, data)
    '3 * x + 8'

    An exception is raised if a variable can't be found. Alternatively, you can also set a default
    value in case this happens.

    >>> _bake_expression('@foo', data, replace_missing_with='bar')
    'bar'

    """

    if callable(expr):
        return expr(data)

    str_expr = _replace_variables(expr, data, replace_missing_with=replace_missing_with)

    # Remove names from named variables
    for pattern in [r"'.+' = ", r"\w+ = "]:
        str_expr = re.sub(pattern, "= ", str_expr)

    return str_expr


def spread(
    template: Template,
    data: Optional[Dict],
    flavor: Flavor,
    postprocess: Optional[Callable] = None,
    start_at: int = 0,
    replace_missing_with: Optional[str] = None,
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
    replace_missing_with
        An optional value to be used when a variable isn't found in the data. An exception is
        raised if a variable is not found and this is not specified.

    Returns
    -------
    cells
        The list of cells.
    n_rows
        The number of rows which the cells span over.

    """

    data = data or {}

    # Unpack the template
    table = []
    for c, col in enumerate(template):
        cells = []
        for r, expr in enumerate(col if isinstance(col, list) else [col]):
            note = None
            if isinstance(expr, tuple):
                expr, note = expr
            cell = _Cell(
                r=r + start_at, c=c, expr=_normalize_expression(expr), note=note
            )
            cells.append(cell)
        table.append(cells)

    # We're going to add the positions of the named variables to the data
    data = data.copy()
    cell_names = {}
    for c, col in enumerate(table):
        for r, cell in enumerate(col):
            if callable(cell.expr):
                cell_names[len(cell_names)] = None
            elif _is_named_formula(cell.expr):
                name = cell.expr.split(" = ")[0]
                data[name] = cell.address
                cell_names[len(cell_names)] = name
            elif _is_variable(cell.expr):
                cell_names[len(cell_names)] = cell.expr[1:]
            else:
                cell_names[len(cell_names)] = None

    if flavor == Flavor.PYGSHEETS.value:
        cells = [
            cell.as_pygsheets(data=data, replace_missing_with=replace_missing_with)
            for col in table
            for cell in col
        ]
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
    replace_missing_with: Optional[str] = None,
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
    replace_missing_with
        An optional value to be used when a variable isn't found in the data. An exception is
        raised if a variable is not found and this is not specified.

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
            replace_missing_with=replace_missing_with,
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
        len(col) if isinstance(col, list) else 1 for col in template.values()
    )

    flat_rows = []

    for k in range(n_rows_in_template, len(df), n_rows_in_template):

        group = df[k - n_rows_in_template : k]

        flat_rows.append(
            {
                var: group[col_name].iloc[i]
                for col_name, col in template.items()
                for i, var in enumerate(col if isinstance(col, list) else [col])
                if var
            }
        )

    return pd.DataFrame(flat_rows)
