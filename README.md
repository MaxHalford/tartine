<h1>🍞 tartine</h1>

- [What this is](#what-this-is)
- [Installation](#installation)
- [Usage example](#usage-example)
  - [Fetching some data](#fetching-some-data)
  - [Getting started](#getting-started)
  - [Adding a header](#adding-a-header)
  - [Linking more cells](#linking-more-cells)
  - [Cell formatting](#cell-formatting)
- [API reference](#api-reference)
- [A note on spreadsheets](#a-note-on-spreadsheets)
- [Development](#development)
- [License](#license)

## What this is

Exporting a dataframe to a spreadsheet is trivial. But this results in a flat and static spreadsheet where the cells are not linked with each other. This is what this tool addresses: it allows you to programmatically generate dynamic spreadsheets with arbitrary layouts.

## Installation

```sh
pip install tartine
```

You can use `tartine` to generate cells with different libraries. You'll have to install these separately. For instance, run `pip install pygsheets` to use [`pygsheets`](https://pygsheets.readthedocs.io/en/stable/index.html).

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

We'll be dumpimp data into a Google Sheet. The [`pygsheets` library](https://pygsheets.readthedocs.io/en/stable/index.html) can be used to interact with Google Sheets.

```py
import pygsheets

gc = pygsheets.authorize(...)
sh = gc.open_by_key('13DneVfUZQlfnKHN2aeo6LUQwCHnjixJ8bV4x092HKqA')
```

You use `tartine` by specifying how you want to spread your data with a template. For this dataset, we want to display the amount of cards per rarity, along with the share each amount represents.

```py
template = [
    "@'Set name'",
    ('Common', 'Rare', 'Epic', 'Legendary'),
    (
        '@Common',
        '@Rare',
        '@Epic',
        '@Legendary',
    ),
    (
        '= @Common / @total',
        '= @Rare / @total',
        '= @Epic / @total',
        '= @Legendary / @total'
    ),
    'total = @Common + @Rare + @Epic + @Legendary'
]
```

This template contains the four different kinds of expressions which `tartine` recognises:

1. `'Common'` is a constant.
2. `@Common` and `@'Set name'` are variables.
3. `= @Common / @total` is a formula.
4. `total = @Common + @Rare + @Epic + @Legendary` is a named formula, which means `@total` can be used elsewhere.

You can generate `pygheets` cells by spreading the data according to the above template:

```py
import tartine

cells = []
nrows = 0

for card_set in card_sets.to_dict('records'):
    _cells, _nrows = tartine.spread(
        template=template,
        data=card_set,
        start_row=nrows,
        flavor='pygsheets'
    )

    cells += _cells
    nrows += _nrows
```

These cells can be sent to the [GSheet](https://docs.google.com/spreadsheets/d/13DneVfUZQlfnKHN2aeo6LUQwCHnjixJ8bV4x092HKqA/edit#gid=0) as so:

```py
wks = sh.worksheet_by_title('v1')
wks.clear()
wks.update_cells(cells)
```

<div align="center">
    <h4><a href="https://docs.google.com/spreadsheets/d/13DneVfUZQlfnKHN2aeo6LUQwCHnjixJ8bV4x092HKqA/edit#gid=0">👀 See the result ✨</a></h4>
</div>

### Adding a header

The GSheet we create in the previous sub-section doesn't contain any column headers. It's trivial to add these, as they're just constants. The tidy thing to do is to turn the template into a dictionary.

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

cells, nrows = tartine.spread(template.keys(), None, flavor='pygsheets')

for card_set in card_sets.to_dict('records'):
    _cells, _nrows = tartine.spread(
        template=template.values(),
        data=card_set,
        start_row=nrows,
        flavor='pygsheets'
    )

    cells += _cells
    nrows += _nrows

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

cells, nrows = tartine.spread(template.keys(), None, flavor='pygsheets')

for card_set in card_sets.to_dict('records'):
    _cells, _nrows = tartine.spread(
        template=template.values(),
        data=card_set,
        start_row=nrows,
        flavor='pygsheets'
    )

    cells += _cells
    nrows += _nrows

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

The `spread` function simply returns a list of `pygsheet.Cell`s, so we can do what we want with them. We can set the `annotate` parameter to `True` to add notes to each cell. This makes it easier to determine what kind of formatting to apply to each cell.

```py
template = {
    'Set name': ("@'Set name'", '', '', ''),
    'Rarity': ('Common', 'Rare', 'Epic', 'Legendary'),
    'Count': (
        'common = @Common',
        'rare = @Rare',
        'epic = @Epic',
        'legendary = @Legendary',
    ),
    'Share': (
        'common_share = @common / @total',
        'rare_share = @rare / @total',
        'epic_share = @epic / @total',
        'legendary_share = @legendary / @total'
    ),
    'Total': ('total = @common + @rare + @epic + @legendary', '', '', '')
}

GRAY = (245 / 255, 245 / 255, 250 / 255, 1)
BLUE = (65 / 255, 105 / 255, 255 / 255, 1)
PURPLE = (191 / 255, 0 / 255, 255 / 255, 1)
ORANGE = (255 / 255, 140 / 255, 0 / 255, 1)

cells, nrows = tartine.spread(template.keys(), None, flavor='pygsheets')
for cell in cells:
    cell.set_text_format('bold', True)

for i, card_set in enumerate(card_sets.to_dict('records')):
    _cells, _nrows = tartine.spread(
        template=template.values(),
        data=card_set,
        start_row=nrows,
        flavor='pygsheets',
        annotate=True
    )

    if i % 2:
        for cell in _cells:
            cell.color = GRAY

    cells += _cells
    nrows += _nrows

for cell in cells:

    if cell.note and 'share' in cell.note:
        cell.set_number_format(
            format_type= pygsheets.FormatType.PERCENT,
            pattern='##0.00%'
        )

    if cell.note and 'rare' in cell.note:
        cell.set_text_format('foregroundColor', BLUE)
    elif cell.note and 'epic' in cell.note:
        cell.set_text_format('foregroundColor', PURPLE)
    elif cell.note and 'legendary' in cell.note:
        cell.set_text_format('foregroundColor', ORANGE)

    cell.note = None

wks = sh.worksheet_by_title('v4')
wks.clear()
wks.update_cells(cells)
```

<div align="center">
    <h4><a href="https://docs.google.com/spreadsheets/d/13DneVfUZQlfnKHN2aeo6LUQwCHnjixJ8bV4x092HKqA/edit#gid=1836554356">👀 See the result ✨</a></h4>
</div>

## API reference

There is a single entrypoint, which is the `spread` function. It has the following parameters:

- `template` — a list of expressions which determines how the cells are layed out.
- `data` — a dictionary of data to render.
- `start_row` — the row number where the layout begins. Zero-based.
- `flavor` — determines what kinds of cells to generate. Only the `pygsheets` flavor is supported right now.
- `annotate` — determines whether or not to attach notes to cells which contain named formulas.

The `spread` function returns the list of cells and the number of rows which the cells span over.

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
pytest
```

## License

The MIT License (MIT). Please see the [license file](LICENSE) for more information.
