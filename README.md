# DeckCraft Fast Skill

[![License: MIT](https://img.shields.io/badge/License-MIT-blue.svg)](LICENSE)
[![Python 3.10+](https://img.shields.io/badge/Python-3.10%2B-3776AB.svg)](https://www.python.org/)
[![DeckCraft](https://img.shields.io/badge/DeckCraft-1--Shot%20Automation-1E40AF.svg)](#)

High-performance presentation generation skill for **DeckCraft**. Eliminates sequential MCP tool roundtrips via 1-shot Python script generation (`User Input -> Plan -> Python Script -> Instant Batch CLI / Live TCP -> Native DeckCraft GUI & PPTX`).

---

## Overview

Traditional AI presentation workflows call MCP tools one by one (`add_slide`, `add_shape`, `set_text`, etc.), causing dozens of LLM turns, high latency, and context window bloat. **DeckCraft Fast Skill** replaces that loop with a structured 5-stage workflow and a lightweight Python automation library (`scripts/`):

```text
[1. User Input] ──▶ [2. Agent Plan] ──▶ [3. User Approval Gateway] ──▶ [4. 1-Shot Python Script] ──▶ [5. Instant Execution (.deckcraft & .pptx + GUI)]
```

### Key Features

- **Hybrid Execution Engine (`DeckCraftClient`)**
  - **Live TCP Streaming**: Connects directly to a running DeckCraft GUI via JSON-RPC 2.0 on `127.0.0.1:7980` (`--control 7980`).
  - **Sub-second Batch CLI**: Falls back to `deckcraft-cli.exe run --cmd-file` and optionally launches `deckcraft.exe` (`auto_gui=True`).
- **Pre-built Design System (`design_tokens.py`)**: Standard 16:9 canvas (`960 x 540 pt`), safe margins (`60 pt` horizontal, `48 pt` vertical), typography scale, and 6 curated color themes.
- **High-Level Slide Components (`DeckComponents`)**: Ready-to-use layout builders (`cover_slide`, `header`, `stat_cards_row`, `feature_cards_row`, `comparison_columns`, `timeline_steps`, `closing_slide`) plus native charts, tables, and vector icons.
- **Dual Output**: Saves native `.deckcraft` files and exports standard Microsoft PowerPoint `.pptx` presentations.

---

## Prerequisites

- **Python 3.10+** (uses Python standard library only; no external `pip` packages required)
- **DeckCraft** installed locally (default Windows paths):
  - `C:\Program Files\DeckCraft\deckcraft-cli.exe`
  - `C:\Program Files\DeckCraft\deckcraft.exe`

---

## Project Structure

```text
deckcraft-fast-SKILL/
├── SKILL.md                        # Agent skill specification & 5-stage Fast Protocol
├── README.md                       # Project documentation
├── LICENSE                         # MIT License
├── .gitignore
├── scripts/
│   ├── deckcraft_ui.py             # DeckCraftClient (Live TCP JSON-RPC & Batch CLI runner)
│   ├── design_tokens.py            # 16:9 canvas constants, typography scale, THEMES, grid helpers
│   └── deckcraft_components.py     # DeckComponents reusable slide layout builders
└── examples/
    └── demo_pitch_deck.py          # End-to-end 5-slide pitch deck demo
```

---

## Quick Start

### 1. Clone the Repository

```bash
git clone https://github.com/muhammad-zidane/deckcraft-fast-SKILL.git
cd deckcraft-fast-SKILL
```

### 2. Run the Example Pitch Deck

```bash
python examples/demo_pitch_deck.py
```

This generates `examples/output/deckcraft_fast_pitch.deckcraft` and `examples/output/deckcraft_fast_pitch.pptx`, then opens the result in the DeckCraft GUI.

### 3. Build a Custom Presentation

```python
import sys
from pathlib import Path

sys.path.insert(0, str(Path("scripts").resolve()))

from deckcraft_ui import DeckCraftClient
from design_tokens import THEMES
from deckcraft_components import DeckComponents

output_dir = Path("./output")
output_dir.mkdir(parents=True, exist_ok=True)

theme = THEMES["corporate_navy"]

with DeckCraftClient(auto_gui=True) as client:
    client.set_outputs(
        save_deckcraft=str(output_dir / "presentation.deckcraft"),
        export_pptx=str(output_dir / "presentation.pptx"),
    )
    comp = DeckComponents(client, theme=theme)

    comp.cover_slide(
        title="DeckCraft Fast AI Engine",
        subtitle="1-shot presentation generation with zero MCP roundtrip latency",
        author="Your Team",
        tag="EXECUTIVE SUMMARY",
    )

    client.new_slide(layout="blank")
    client.set_background(theme.bg_canvas)
    comp.header(title="Key Metrics", subtitle="Performance overview", category="BENCHMARKS")
    comp.stat_cards_row([
        {"value": "98%", "label": "Time Saved", "subtext": "vs sequential MCP calls", "icon": "zap"},
        {"value": "<1s", "label": "Render Speed", "subtext": "Native CLI execution", "icon": "trending-up"},
    ])

    comp.closing_slide(
        title="Thank You",
        subtitle="Ready to automate your presentation workflow.",
        cta_text="Get Started",
    )
```

---

## Design System & Components

### Available Themes (`THEMES`)

| Key | Name | Best For |
| :--- | :--- | :--- |
| `corporate_navy` | Corporate Navy | Business presentations, investor decks, executive reports |
| `modern_dark` | Modern Dark | Tech products, AI, developer tooling |
| `clean_slate` | Clean Slate | Modern startup pitch decks |
| `emerald_growth` | Emerald Growth | Finance, sustainability, sales |
| `vibrant_purple` | Vibrant Purple | Creative & innovation showcases |
| `warm_editorial` | Warm Editorial | Case studies, brand storytelling, education |

### Canvas Specifications (`design_tokens.py`)

- **Aspect Ratio / Size**: `16:9` (`960 x 540 pt`)
- **Safe Margins**: Left/Right `60 pt`, Top/Bottom `48 pt` (`840 x 444 pt` effective content area)
- **Standard Y Coordinates**: Header `48 pt`, Subtitle `96 pt`, Body `150 pt` (height `330 pt`), Footer `500 pt`

### Reusable Components (`DeckComponents`)

- `cover_slide(title, subtitle, author=None, tag="PRESENTATION", date_str=None)`
- `header(title, subtitle=None, category=None)`
- `stat_cards_row(stats, y=160, height=170)`
- `feature_cards_row(cards, y=150, height=320)`
- `comparison_columns(left_col, right_col, y=150, height=320)`
- `timeline_steps(steps, y=200)`
- `closing_slide(title, subtitle, cta_text, contact_info=None)`

For advanced primitives (`client.add_chart`, `client.add_table`, `client.add_icon`), see [SKILL.md](SKILL.md).

---

## Using as an AI Agent Skill

Point your AI coding assistant to [`SKILL.md`](SKILL.md) (or install this repository under your agent's skills directory, e.g. `~/.gemini/config/skills/deckcraft-fast-skill`). The skill enforces a **User Approval Gateway** after the slide plan is drafted before any Python script is generated or executed.

---

## License

Distributed under the MIT License. See [LICENSE](LICENSE) for more information.
