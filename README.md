# Caseboard

A solo, Wordle-style daily deduction game built on Codenames mechanics: a fixed 5x5 word
grid and a scripted clue sequence that's identical for every player, scored on how few
clues and how few wrong guesses it takes to clear the board.

This is a working prototype (not yet a real daily-rotation game) — one hand-authored board,
playable entirely client-side, no backend. Built with Claude.

## Play it

Live at https://froobaloo.com/caseboard/, part of the Froobaloo games hub. The
`jtoeman/froobaloo` repo copies these files from `main` and deploys them; every push
here asks it to rebuild (needs the `HUB_TOKEN` secret). This repo's own GitHub Pages
address only redirects to froobaloo.com. You can also open `index.html` directly.

## How it works

- Three opening clues (category + count) start the case, covering every target word.
- Any target word not found by the opening clues gets its own one-word follow-up clue —
  skipped automatically for anything already found.
- No hard fail: if you run out of scripted clues, the board stays open for free guessing
  until every target is found.
- Two words on the board are double-cost — guessing one wrong counts as two misses.
- Score = clues used + 1 per miss + 2 per double-cost miss. Lower is better, par is the
  number of opening clues.
