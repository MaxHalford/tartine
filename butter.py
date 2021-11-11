import re
import string
from collections import UserDict, UserList
from dataclasses import dataclass
from typing import Dict, Tuple

import pygsheets


def column_letter(n: int) -> str:
    code = ""
    n += 1
    while n > 0:
        n, mod = divmod(n - 1, 26)
        code = chr(65 + mod) + code
    return code


def is_variable(expr):
    return isinstance(expr, str) and expr.startswith("@")


def is_formula(expr):
    return isinstance(expr, str) and expr.startswith("=")


def is_named_formula(expr):
    return isinstance(expr, str) and bool(re.match(r"@[\w_\.]+ = ", expr))


def normalize_expression(expr):
    return re.sub(r"\s+", " ", expr).strip()


def make_layout(template):

    layout = {}

    for j, expr in enumerate(template.values()):

        if isinstance(expr, str):
            expr = [expr]

        for i, e in enumerate(expr):
            layout[i, j] = normalize_expression(e)

    return layout


def get_nested(d, path):
    result = d
    for part in path.split("."):
        result = result[part]
    return result


@dataclass
class Cell:
    row: int
    column: int
    expression: str


class Cells(UserDict):
    def __init__(
        self,
        template,
    ):
        super().__init__()
        for j, (name, expr) in enumerate(template.items()):
            self[name] = []
            for i, e in enumerate([expr] if isinstance(expr, str) else expr):
                cell = Cell(row=i, column=j, expression=normalize_expression(e))
                self[name].append(cell)

    @property
    def n_rows(self):
        return max(map(len, self.values()))

    def make_pygsheets(self, instance, start_row, formatting=None):

        if formatting is None:
            formatting = {}

        # Column name -> A0
        absolute_positions = {
            name: [f"{column_letter(cell.column)}{start_row + cell.row + 1}" for cell in cells]
            for name, cells in self.items()
        }

        # @foo -> A0
        variable_absolute_positions = {
            (
                re.split(" = ", self[name][i].expression)[0]
                if is_named_formula(self[name][i].expression)
                else self[name][i].expression
            ): position
            for name, positions in absolute_positions.items()
            for i, position in enumerate(positions)
            if is_variable(self[name][i].expression)
        }

        cell_values = {}

        for name, cells in self.items():
            cell_values[name] = []
            for cell in cells:
                expr = cell.expression

                if is_named_formula(expr):
                    expr = expr.split(" ", 1)[1]

                if is_formula(expr):
                    for var in re.findall(r"@[\w_\.]+", expr):
                        if var in variable_absolute_positions:
                            expr = expr.replace(var, variable_absolute_positions[var])
                        else:
                            expr = expr.replace(var, str(get_nested(instance, var.replace("@", ""))))

                cell_values[name].append(expr)

        # Replace variables with the values in the instance
        for name, expressions in cell_values.items():
            for i, expr in enumerate(expressions):
                if is_variable(expr):
                    cell_values[name][i] = get_nested(instance, expr[1:])

        gcells = []
        for name, cells in self.items():
            for i, cell in enumerate(cells):
                value = cell_values[name][i]
                gcell = pygsheets.Cell((start_row + cell.row + 1, cell.column + 1), value)
                if is_formula(value):
                    gcell.formula = value

                cell_formatting = formatting.get(name)
                if cell_formatting:
                    gcell = gcell.set_number_format(
                        format_type=cell_formatting["type"],
                        pattern=cell_formatting["pattern"],
                    )
                gcells.append(gcell)

        return gcells
