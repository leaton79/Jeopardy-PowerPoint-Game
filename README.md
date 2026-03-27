# Excel to Jeopardy `.pptm`

Generate an interactive Jeopardy-style PowerPoint from a spreadsheet.

This project exists because a plain `.pptx` cannot maintain true cumulative arbitrary-order board state during slideshow runtime. To make the board clear clues permanently in any order, the final output must be a macro-enabled `.pptm`.

## What it does

- reads clue data from a spreadsheet
- builds a Jeopardy board slide, clue slides, answer slides, and a final `Game Over` slide
- injects VBA so the board state updates cumulatively during Slide Show mode
- assigns PowerPoint click actions so the flow is:

`board tile -> clue slide -> answer slide -> board slide`

## Requirements

- macOS
- Microsoft PowerPoint for Mac
- Python 3.11+
- `pandas`
- `python-pptx`
- `openpyxl`

## Input format

The generator expects a workbook with columns matching these logical fields:

- `Category`
- `Value`
- `Clue`
- `Answer`

The exact header names may vary slightly. The script detects close matches automatically.

## Project layout

- `generate_jeopardy_pptm.py`: main generator
- `sample_data/sample_jeopardy_board.xlsx`: example source workbook
- `examples/jeopardy_game_example.pptm`: example generated deck
- `requirements.txt`: Python dependencies

## Setup

```bash
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
```

## Usage

```bash
python3 generate_jeopardy_pptm.py \
  --input sample_data/sample_jeopardy_board.xlsx \
  --output build/jeopardy_game.pptm \
  --report build/jeopardy_macro_report.txt
```

## Running the slideshow

1. Open the generated `.pptm` in Microsoft PowerPoint for Mac.
2. If PowerPoint shows a security prompt, enable macros.
3. Start Slide Show or Presenter Mode.
4. Click a board tile to open a clue.
5. Click `Reveal Answer` to move from the clue slide to the answer slide.
6. Click `Return to Board` to mark that clue as used and go back to the main board.
7. Continue in any order until the deck routes to `Game Over`.

The cumulative board state is handled by VBA during the slideshow. Without macros enabled, the deck will open, but the board will not update correctly as clues are played.

## How it works

The Python layer builds the deck structure:

- slide 1: board
- slides 2..N: clue/answer pairs
- final slide: `Game Over`

The VBA layer maintains runtime state:

- each clue gets a stable ID
- a `usedFlags()` array tracks which clues have been played
- clicking a board tile runs a `GoToClue_*` macro
- clicking the clue button runs a `Reveal_*` macro
- clicking the answer button runs a `Complete_*` macro
- each `Complete_*` macro marks that clue as used, refreshes the board, and returns to slide 1
- when all clues are used, VBA jumps to `Game Over`

## How to adapt it for a new game

1. Replace the workbook with your own spreadsheet.
2. Keep one row per clue.
3. Make sure each row has a category, point value, clue text, and answer text.
4. Avoid duplicate category/value pairs within the same board.
5. Run the generator again with the new input file.

If the workbook uses slightly different column names such as `Question` instead of `Clue`, the script attempts to match those automatically. If it cannot find the required fields, it fails with a clear error.

## Expected deck behavior

- The board can be played in arbitrary order.
- Once a clue is completed, that tile stays cleared.
- Remaining clues stay clickable.
- The last completed clue sends the slideshow to `Game Over`.

This behavior is implemented with PowerPoint macros, not with static hyperlink-only slide duplication.

## Important limitation

This project depends on PowerPoint automation to create the final macro-enabled file. `python-pptx` can build the slide content, but it cannot create a VBA project by itself.

## Publishing notes

Before pushing to GitHub, review the example deck and sample workbook to make sure they do not contain private or licensed material you do not want to publish.

## License

GNU GPL v3.0
