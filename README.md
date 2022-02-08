<h1>🍞 tartine</h1>

- [What this is](#what-this-is)
- [Installation](#installation)
- [Usage example](#usage-example)
  - [Fetching some data](#fetching-some-data)
  - [Getting started](#getting-started)
  - [Spreading a dataframe](#spreading-a-dataframe)
  - [Linking more cells](#linking-more-cells)
  - [Cell formatting](#cell-formatting)
- [API reference](#api-reference)
  - [`spread`](#spread)
  - [`spread_dataframe`](#spread_dataframe)
- [Supported flavors](#supported-flavors)
- [A note on spreadsheets](#a-note-on-spreadsheets)
- [Development](#development)
- [License](#license)

## What this is

Exporting a dataframe to a spreadsheet is trivial. But this results in a flat and static spreadsheet where the cells are not linked with each other. This is what this tool addresses: it allows you to programmatically generate dynamic spreadsheets with arbitrary layouts.

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

print(card_sets.head())
```

| Set name                                          | Release type   | Release date      | Removal date from Standard   |   Total |   Common |   Rare |   Epic |   Legendary |
|:--------------------------------------------------|:---------------|:------------------|:-----------------------------|--------:|---------:|-------:|-------:|------------:|
| United in Stormwind                               | Expansion      | August 3, 2021    | TBA 2023                     |     135 |       50 |     35 |     25 |          25 |
| Forged in the Barrens with Wailing Caverns        | Expansion      | March 30, 2021    | TBA 2023                     |     170 |       66 |     49 |     26 |          29 |
| Core 2021 (Core)                                  | Core           | March 30, 2021    | TBA 2022                     |     235 |      128 |     55 |     27 |          25 |
| Madness at the Darkmoon Faire with Darkmoon Races | Expansion      | November 17, 2020 | TBA 2022                     |     170 |       70 |     46 |     25 |          29 |
| Scholomance Academy (Scholomance)                 | Expansion      | August 6, 2020    | TBA 2022                     |     135 |       52 |     35 |     23 |          25 |

### Getting started

You use `tartine` by specifying how you want to spread your data with a template. For this dataset, we want to display the amount of cards per rarity, along with the share each amount represents.

```py
template = {
    'Set name': "@'Set name'",
    'Rarity': ('Common', 'Rare', 'Epic', 'Legendary'),
    'Count': (
        '@Common',
        '@Rare',
        '@Epic',
        '@Legendary',
    ),
    'Share': (
        '= @Common / @total',
        '= @Rare / @total',
        '= @Epic / @total',
        '= @Legendary / @total'
    ),
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
wks.clear()
wks.update_cells(cells)
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
wks.clear()
wks.update_cells(cells)
```

<div align="center">
    <h4><a href="https://docs.google.com/spreadsheets/d/13DneVfUZQlfnKHN2aeo6LUQwCHnjixJ8bV4x092HKqA/edit#gid=709697806">👀 See the result ✨</a></h4>
</div>

### Linking more cells

The spreadsheet we built displays the data in a static manner. The share of each rarity is obtained by dividing the amount of cards by the total amount. You'll notice that the total doesn't change if you manually edit any of the amounts. This is because it's calculated from the data, and isn't actually referencing any of the cells. We can change this by using named formulas instead of variables.

```py
template = {
    'Set name': "@'Set name'",
    'Rarity': ('Common', 'Rare', 'Epic', 'Legendary'),
    'Count': (
        'common = @Common',
        'rare = @Rare',
        'epic = @Epic',
        'legendary = @Legendary',
    ),
    'Share': (
        '= @common / @total',
        '= @rare / @total',
        '= @epic / @total',
        '= @legendary / @total'
    ),
    'Total': 'total = @common + @rare + @epic + @legendary'
}

cells = tartine.spread_dataframe(
    template=template,
    df=card_sets,
    flavor='pygsheets'
)

wks = sh.worksheet_by_title('v3')
wks.clear()
wks.update_cells(cells)
```

Now you should see the cell values update automatically when you modify any of the amounts.

<div align="center">
    <h4><a href="https://docs.google.com/spreadsheets/d/13DneVfUZQlfnKHN2aeo6LUQwCHnjixJ8bV4x092HKqA/edit#gid=2042929262">👀 See the result ✨</a></h4>
</div>

### Cell formatting

The sheet we have displays the data correctly and the cells are linked with each other. Yipee. However, it's a bit ugly, and it would be nice to also format the cells programmatically. Indeed, readability would be improved by adding some colors and formatting the percentages.

First of all there is a `postprocess` parameter that allows to do any kind of transformation to each cell once it has been created. This can be used to pass a `stylize` function which applies the adequate modifications.

Secondly you can do whatever you want to the returned cells. Here we will check the row number of all the cells and shade them accordingly. We'll also bolden the font of the cells in first row.

```py
GRAY = (245 / 255, 245 / 255, 250 / 255, 1)
BLUE = (65 / 255, 105 / 255, 255 / 255, 1)
PURPLE = (191 / 255, 0 / 255, 255 / 255, 1)
ORANGE = (255 / 255, 140 / 255, 0 / 255, 1)

def stylize(cell, name):

    # Format percentages
    if name and 'share' in name:
        cell.set_number_format(
            format_type= pygsheets.FormatType.PERCENT,
            pattern='##0.00%'
        )

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

for cell in cells:
    # Bolden the header
    if cell.row == 1:
        cell.set_text_format('bold', True)
    # Shade every group of 4 rows
    if any(cell.row % 8 - r == 0 for r in (2, 3, 4, 5)):
        cell.color = GRAY

wks = sh.worksheet_by_title('v4')
wks.clear()
wks.update_cells(cells)
```

<div align="center">
    <h4><a href="https://docs.google.com/spreadsheets/d/13DneVfUZQlfnKHN2aeo6LUQwCHnjixJ8bV4x092HKqA/edit#gid=1836554356">👀 See the result ✨</a></h4>
</div>

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
        A dictionary of data to render.
    flavor
        Determines what kind of cells to generate.
    postprocess
        An optional function to call for each cell once it has been created.
    start_at
        The row number where the layout begins. Zero-based.
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
    df
        A dataframe to render.
    template
        A list of expressions which determines how the cells are layed out.
    flavor
        Determines what kind of cells to generate.
    postprocess
        An optional function to call for each cell once it has been created.
<BLANKLINE>
    Returns
    -------
    cells
        The list of cells.
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

I always thought spreadsheets sucked. I still think they do in many cases. But they definitely fit the bill for some tasks. This was very much true when I worked at [Alan](https://alan.com/). They really shine when you need to build an interactive app, but can't afford to spend engineering resources. One might even go as far to say that they're underrated. Let me quote [this](https://news.ycombinator.com/item?id=29104047#29108603) Hackernews comment:

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
