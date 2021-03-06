

<div align="center">
  <h1>🍞 tartine</h1>
  <p>Manipulate dynamic spreadsheets with arbitrary layouts using Python.</p>
</div>
<br>

<div align="center">
  <!-- Tests -->
  <a href="https://github.com/MaxHalford/tartine/actions/workflows/unit-tests.yml">
    <img src="https://github.com/MaxHalford/tartine/actions/workflows/unit-tests.yml/badge.svg?style=flat-square" alt="unit-tests">
  </a>
  <!-- PyPI -->
  <a href="https://pypi.org/project/tartine">
    <img src="https://img.shields.io/pypi/v/tartine.svg?label=release&color=blue&style=flat-square" alt="pypi">
  </a>
  <!-- License -->
  <a href="https://opensource.org/licenses/MIT">
    <img src="https://img.shields.io/badge/License-MIT-blue.svg?style=flat-square" alt="license">
  </a>
</div>
<br>

- [What this is](#what-this-is)
- [Installation](#installation)
- [Usage example](#usage-example)
  - [Fetching some data](#fetching-some-data)
  - [Getting started](#getting-started)
  - [Spreading a dataframe](#spreading-a-dataframe)
  - [Linking cells](#linking-cells)
  - [Cell styling](#cell-styling)
  - [Styling one cell at a time](#styling-one-cell-at-a-time)
  - [Adding notes](#adding-notes)
  - [Unspreading a dataframe](#unspreading-a-dataframe)
  - [Handling nested data](#handling-nested-data)
- [API reference](#api-reference)
  - [`spread`](#spread)
  - [`spread_dataframe`](#spread_dataframe)
  - [`unspread_dataframe`](#unspread_dataframe)
- [Supported flavors](#supported-flavors)
- [A note on spreadsheets](#a-note-on-spreadsheets)
- [Development](#development)
- [License](#license)

## What this is

Exporting a dataframe to a spreadsheet is trivial. But this results in a flat and static spreadsheet where the cells are not linked with each other. This is what this tool addresses: it allows you to programmatically generate dynamic spreadsheets with arbitrary layouts.

This tool also allows to do things the other way round: convert a dataframe with an arbitrary layout into a flat dataframe.

## Installation

```sh
pip install tartine
```

You can use `tartine` to generate cells with different libraries. You'll have to install these separately. For instance, run `pip install pygsheets` to use [`pygsheets`](https://pygsheets.readthedocs.io/en/stable/index.html). The [`pandas`](https://pandas.pydata.org/) library is also not installed by default.

## Usage example

### Fetching some data

Hearthstone is a virtual card game which I'll use as an example. New card sets are released every so often. The list of card sets is available [on Wikipedia](https://en.wikipedia.org/wiki/Hearthstone?oldformat=true#Card_sets). It's straightforward to fetch this data with a little bit of pandas judo.

```py
import pandas as pd

card_sets = pd.read_html(
    'https://en.wikipedia.org/wiki/Hearthstone#Card_sets',
    match='Collectible cards breakdown',
)[0]

card_sets = card_sets.rename(columns={
    'Set name (abbreviation)': 'Set name',
    'Removal datefrom Standard': 'Removal date from Standard'
})
card_sets = card_sets[~card_sets['Set name'].str.startswith('Year')]
for col in ['Set name', 'Release date']:
    card_sets[col] = card_sets[col].str.replace(r'\[.+\]', '', regex=True)
card_sets = card_sets[1:-1]
card_sets = card_sets[::-1]  # latest to oldest

print(card_sets.head().to_markdown(index=False))
```

| Set name                                          | Release type   | Release date      | Removal date from Standard   |   Total |   Common |   Rare |   Epic |   Legendary |
|:--------------------------------------------------|:---------------|:------------------|:-----------------------------|--------:|---------:|-------:|-------:|------------:|
| Fractured in Alterac Valley                       | Expansion      | December 7, 2021  | TBA 2023                     |     135 |       50 |     35 |     24 |          26 |
| United in Stormwind with Deadmines                | Expansion      | August 3, 2021    | TBA 2023                     |     170 |       66 |     49 |     26 |          29 |
| Forged in the Barrens with Wailing Caverns        | Expansion      | March 30, 2021    | TBA 2023                     |     170 |       66 |     49 |     26 |          29 |
| Core 2021 (Core)                                  | Core           | March 30, 2021    | TBA 2022                     |     235 |      128 |     55 |     27 |          25 |
| Madness at the Darkmoon Faire with Darkmoon Races | Expansion      | November 17, 2020 | TBA 2022                     |     170 |       70 |     46 |     25 |          29 |

### Getting started

You use `tartine` by specifying how you want to spread your data with a template. For this dataset, we want to display the amount of cards per rarity, along with the share each amount represents.

```py
template = {
    'Set name': '@Set name',
    'Rarity': ['Common', 'Rare', 'Epic', 'Legendary'],
    'Count': [
        '@Common',
        '@Rare',
        '@Epic',
        '@Legendary',
    ],
    'Share': [
        '= @Common / @total',
        '= @Rare / @total',
        '= @Epic / @total',
        '= @Legendary / @total'
    ],
    'Total': 'total = @Common + @Rare + @Epic + @Legendary'
}
```

This template contains the four different kinds of expressions which `tartine` recognises:

1. `'Common'` is a constant.
2. `@Common` is a variable.
3. `= @Common / @total` is a formula.
4. `total = @Common + @Rare + @Epic + @Legendary` is a named formula, which means `@total` can be used elsewhere.

You can generate `pygsheets.Cell`s by spreading the data according to the above template:

```py
import tartine

cells = []
nrows = 0

for card_set in card_sets.to_dict('records'):
    _cells, _nrows = tartine.spread(
        template=template.values(),
        data=card_set,
        start_at=nrows,
        flavor='pygsheets'
    )

    cells += _cells
    nrows += _nrows
```

We'll dump data into [this](https://docs.google.com/spreadsheets/d/13DneVfUZQlfnKHN2aeo6LUQwCHnjixJ8bV4x092HKqA) public Google Sheet for the sake of example. The [`pygsheets` library](https://pygsheets.readthedocs.io/en/stable/index.html) can be used to interact with Google Sheets.

```py
import pygsheets

gc = pygsheets.authorize(...)
sh = gc.open_by_key('13DneVfUZQlfnKHN2aeo6LUQwCHnjixJ8bV4x092HKqA')
```

These cells can be sent to the [GSheet](https://docs.google.com/spreadsheets/d/13DneVfUZQlfnKHN2aeo6LUQwCHnjixJ8bV4x092HKqA/edit#gid=0) like so:

```py
wks = sh.worksheet_by_title('v1')
wks.clear(fields='*')
wks.update_values(cell_list=cells)
```

<div align="center">
    <h4><a href="https://docs.google.com/spreadsheets/d/13DneVfUZQlfnKHN2aeo6LUQwCHnjixJ8bV4x092HKqA/edit#gid=0">👀 See the result ✨</a></h4>
</div>

### Spreading a dataframe

What we just did was a bit manual. We had to loop through the rows of the dataframe and concatenate the cells ourselves. On the one hand that gives you a lot of freedom. On the other hand you'll probably be working with `pandas.DataFrame`s in practice, so you'll want to avoid this kind of boilerplate.

The `spread_dataframe` allows you to do what we just did with a one-liner. As a bonus the column names are included.

```py
cells = tartine.spread_dataframe(
    template=template,
    df=card_sets,
    flavor='pygsheets'
)

wks = sh.worksheet_by_title('v2')
wks.clear(fields='*')
wks.update_values(cell_list=cells)
```

<div align="center">
    <h4><a href="https://docs.google.com/spreadsheets/d/13DneVfUZQlfnKHN2aeo6LUQwCHnjixJ8bV4x092HKqA/edit#gid=709697806">👀 See the result ✨</a></h4>
</div>

### Linking cells

The spreadsheet we built displays the data in a static manner. The share of each rarity is obtained by dividing the amount of cards by the total amount. You'll notice that the total doesn't change if you manually edit any of the amounts. This is because it's calculated from the data, and isn't actually referencing any of the cells. We can change this by using named formulas instead of variables.

```py
template.update({
    'Count': [
        'common = @Common',
        'rare = @Rare',
        'epic = @Epic',
        'legendary = @Legendary',
    ],
    'Share': [
        '= @common / @total',
        '= @rare / @total',
        '= @epic / @total',
        '= @legendary / @total'
    ],
    'Total': 'total = @common + @rare + @epic + @legendary'
})

cells = tartine.spread_dataframe(
    template=template,
    df=card_sets,
    flavor='pygsheets'
)

wks = sh.worksheet_by_title('v3')
wks.clear(fields='*')
wks.update_values(cell_list=cells)
```

Now you should see the cell values update automatically when you modify any of the amounts.

<div align="center">
    <h4><a href="https://docs.google.com/spreadsheets/d/13DneVfUZQlfnKHN2aeo6LUQwCHnjixJ8bV4x092HKqA/edit#gid=2042929262">👀 See the result ✨</a></h4>
</div>

### Cell styling

The sheet we have displays the data correctly and the cells are linked with each other. Yipee. However, it's a bit ugly, and it would be nice to also format the cells programmatically. Indeed, readability would be improved by adding some colors.

There is a `postprocess` parameter that allows to do any kind of transformation to each cell once it has been created. This can be used to pass a `stylize` function which applies the adequate modifications.

```py
# Let's add empty for background coloring
template['Set name'] = ['@Set name', '', '', '']
template['Total'] = ['total = @common + @rare + @epic + @legendary', '', '', '']

GRAY = (245 / 255, 245 / 255, 250 / 255, 1)
BLUE = (65 / 255, 105 / 255, 255 / 255, 1)
PURPLE = (191 / 255, 0 / 255, 255 / 255, 1)
ORANGE = (255 / 255, 140 / 255, 0 / 255, 1)

def stylize(cell, name):

    # Bolden the header
    if cell.row == 1:
        cell.set_text_format('bold', True)

    # Shade every group of 4 rows
    if any(cell.row % 8 - r == 0 for r in (2, 3, 4, 5)):
        cell.color = GRAY

    # Color by rarity
    if name and 'rare' in name:
        cell.set_text_format('foregroundColor', BLUE)
    elif name and 'epic' in name:
        cell.set_text_format('foregroundColor', PURPLE)
    elif name and 'legendary' in name:
        cell.set_text_format('foregroundColor', ORANGE)

    return cell

cells = tartine.spread_dataframe(
    template=template,
    df=card_sets,
    flavor='pygsheets',
    postprocess=stylize
)

wks = sh.worksheet_by_title('v4')
wks.clear(fields='*')
wks.update_values(cell_list=cells)
```

<div align="center">
    <h4><a href="https://docs.google.com/spreadsheets/d/13DneVfUZQlfnKHN2aeo6LUQwCHnjixJ8bV4x092HKqA/edit#gid=1836554356">👀 See the result ✨</a></h4>
</div>

### Styling one cell at a time

The `postprocess` argument allows you to style cells in a global manner. You can also pass an extra value to a template entry to style a cell in particular. As an example, let's format the percentages.

```py
def format_pct(cell):
    cell.set_number_format(
        format_type=pygsheets.FormatType.PERCENT,
        pattern='0%'
    )
    return cell

template['Share'] = [
    ('= @common / @total', format_pct),
    ('= @rare / @total', format_pct),
    ('= @epic / @total', format_pct),
    ('= @legendary / @total', format_pct)
]

cells = tartine.spread_dataframe(
    template=template,
    df=card_sets,
    flavor='pygsheets',
    postprocess=stylize
)

wks = sh.worksheet_by_title('v5')
wks.clear(fields='*')
wks.update_values(cell_list=cells)
```

<div align="center">
    <h4><a href="https://docs.google.com/spreadsheets/d/13DneVfUZQlfnKHN2aeo6LUQwCHnjixJ8bV4x092HKqA/edit#gid=2046920632">👀 See the result ✨</a></h4>
</div>

### Adding notes

You may also want to include cell notes. This can be done by padding yet another argument to the relevant template entries. For instance, let's add a note detailing how each share is calculated.

```py
template['Share'] = [
    (
        '= @common / @total',
        format_pct,
        '@Common / (@Common + @Rare + @Epic + @Legendary)'
    ),
    (
        '= @rare / @total',
        format_pct,
        '@Rare / (@Common + @Rare + @Epic + @Legendary)'
    ),
    (
        '= @epic / @total',
        format_pct,
        '@Epic / (@Common + @Rare + @Epic + @Legendary)'
    ),
    (
        '= @legendary / @total',
        format_pct,
        '@Legendary / (@Common + @Rare + @Epic + @Legendary)'
    )
]

cells = tartine.spread_dataframe(
    template=template,
    df=card_sets,
    flavor='pygsheets',
    postprocess=stylize
)

wks = sh.worksheet_by_title('v6')
wks.clear(fields='*')
wks.update_values(cell_list=cells)
```

<div align="center">
    <h4><a href="https://docs.google.com/spreadsheets/d/13DneVfUZQlfnKHN2aeo6LUQwCHnjixJ8bV4x092HKqA/edit#gid=2111260219">👀 See the result ✨</a></h4>
</div>

### Unspreading a dataframe

There's also way to "unspread" a structured sheet into a flat dataframe. We'll take the last version we created as an example.

```py
v4 = sh.worksheet_by_title('v4').get_as_df(start='A1')
print(v4.head(8).to_markdown(index=False))
```

| Set name                           | Rarity    |   Count | Share   | Total   |
|:-----------------------------------|:----------|--------:|:--------|:--------|
| Fractured in Alterac Valley        | Common    |      50 | 37.04%  | 135     |
|                                    | Rare      |      35 | 25.93%  |         |
|                                    | Epic      |      24 | 17.78%  |         |
|                                    | Legendary |      26 | 19.26%  |         |
| United in Stormwind with Deadmines | Common    |      66 | 38.82%  | 170     |
|                                    | Rare      |      49 | 28.82%  |         |
|                                    | Epic      |      26 | 15.29%  |         |
|                                    | Legendary |      29 | 17.06%  |         |


This dataframe can be flattened by providing a template to the `unspread_dataframe` pattern. This keys of this template are the column names of the input dataframe. The values are sequences with variable names. Each variable will correspond to a column in the output dataframe.

```py
template = {
    'Set name': ['set_name', '', '', ''],
    'Total': 'total',  # same, but shorter
    'Count': ['common', 'rare', 'epic', 'legendary']
}

v4_flat = tartine.unspread_dataframe(template, v4)
print(v4_flat.head(2).to_markdown(index=False))
```

| set_name                           |   total |   common |   rare |   epic |   legendary |
|:-----------------------------------|--------:|---------:|-------:|-------:|------------:|
| Fractured in Alterac Valley        |     135 |       50 |     35 |     24 |          26 |
| United in Stormwind with Deadmines |     170 |       66 |     49 |     26 |          29 |

### Handling nested data

Your data might be nested when you're using the `spread` function. For instance:

```py
data = {
    'rarity': {
        'common': 50,
        'rare cards': 35,
        'epic': 24,
        'legendary': 26
    }
}
```

Under the hood, `tartine` uses the [`glom` library](https://glom.readthedocs.io/en/latest/). This means you can use dotted expressions, as so:

```py
template = {
    'Rarity': ['Common', 'Rare', 'Epic', 'Legendary'],
    'Count': [
        'common = @rarity.common',
        'rare = @rarity.rare cards',  # note the space
        'epic = @rarity.epic',
        'legendary = @rarity.legendary',
    ],
    'Share': [
        '= @common / @total',
        '= @rare / @total',
        '= @epic / @total',
        '= @legendary / @total'
    ],
    'Total': 'total = @common + @rare + @epic + @legendary'
}

cells, nrows = tartine.spread(template.values(), data, flavor='pygsheets')
```

In case string expressions are not enough, you can also pass a function which takes as input the data.

```py
template = {
    'Total': lambda data: (
        data['rarity']['common'] +
        data['rarity']['rare'] +
        data['rarity']['epic'] +
        data['rarity']['legendary']
    }
}

cells, nrows = tartine.spread(template.values(), data, flavor='pygsheets')
```

## API reference

### `spread`

```py
>>> import tartine
>>> print(tartine.spread.__doc__)
Spread data into cells.
<BLANKLINE>
    Parameters
    ----------
    template
        A list of expressions which determines how the cells are layed out.
    data
        Data to render. Can be a dictionary, a dataclass, a list; just as long as the template
        expressions can be applied to the data.
    flavor
        Determines what kind of cells to generate.
    postprocess
        An optional function to call for each cell once it has been created.
    start_at
        The row number where the layout begins. Zero-based.
    replace_missing_with
        An optional value to be used when a variable isn't found in the data. An exception is
        raised if a variable is not found and this is not specified.
<BLANKLINE>
    Returns
    -------
    cells
        The list of cells.
    n_rows
        The number of rows which the cells span over.
<BLANKLINE>
<BLANKLINE>

```

### `spread_dataframe`

```py
>>> print(tartine.spread_dataframe.__doc__)
Spread a dataframe into cells.
<BLANKLINE>
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
<BLANKLINE>
    Returns
    -------
    cells
        The list of cells.
<BLANKLINE>
<BLANKLINE>

```

### `unspread_dataframe`

```py
>>> print(tartine.unspread_dataframe.__doc__)
Unspread a dataframe into a flat dataframe.
<BLANKLINE>
    Parameters
    ----------
    template
        A list of expressions which determines how the cells are layed out.
    df
        A dataframe to unspread. Typically this may be the output of the `spread` function once
        it has been dumped into a sheet.
<BLANKLINE>
    Returns
    -------
    flat_df
        The flattened dataframe.
<BLANKLINE>
<BLANKLINE>

```

## Supported flavors

The `spread` and `spread_dataframe` methods accept a `flavor` parameter. It determines what kind of library to use for producing cells.

```py
>>> for flavor in tartine.Flavor:
...     print(flavor.value)
pygsheets

```

## A note on spreadsheets

I always thought spreadsheets sucked. I still think they do in many cases. But they definitely fit the bill for some tasks. This was very much true when I worked at [Alan](https://alan.com/). Likewise at [Carbonfact](https://www.carbonfact.com/). They really shine when you need to build an interactive app, but can't afford to spend engineering resources. One might even go as far to say that they're underrated. Let me quote [this](https://news.ycombinator.com/item?id=29104047#29108603) Hackernews comment:

> > I've mentioned that programmers are far too dismissive of MS Excel. You can achieve a awful lot with Excel: more, even, than some programmers can achieve without it
>
> This is one of the most underrated topics in tech imho. Spreadsheet is probably the pinnacle of how tech could be easily approachable by non tech people, in the "bike for the mind" sense. We came a long way down hill from there when you need an specialist even to come up with a no-code solution to mundane problems.
>
> Sure the tech ecosystem evolved and became a lot more complex from there but I'm afraid the concept of a non-tech person opening a blank file and creating something useful from scratch has been lost along the way.

I also like the idea that spreadsheets can convey information through their layout. For instance, this [banana nut bread recipe](https://twitter.com/craigmod/status/1458289159479586821/photo/1) is a great example of data literacy. It's definitely something you can build with `tartine`.

## Development

```sh
git clone https://github.com/MaxHalford/tartine
cd tartine
pip install poetry
poetry install
poetry shell
pytest
```

## License

The MIT License (MIT). Please see the [license file](LICENSE) for more information.
