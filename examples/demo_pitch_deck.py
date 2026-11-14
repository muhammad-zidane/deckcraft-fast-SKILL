"""
demo_pitch_deck.py - End-to-end 5-slide presentation demo using deckcraft-fast-SKILL.
Demonstrates:
  1. Instant 1-shot batch generation (0 MCP round-trips).
  2. Professional theme design tokens.
  3. Reusable presentation components (Cover, Problem/Solution, KPIs, Workflow, CTA).
  4. Automatic export to .deckcraft and .pptx.
  5. Native GUI display.
"""

import sys
from pathlib import Path

# Add scripts directory to sys.path
SCRIPTS_DIR = Path(__file__).resolve().parent.parent / "scripts"
sys.path.insert(0, str(SCRIPTS_DIR))

from deckcraft_ui import DeckCraftClient
from design_tokens import THEMES
from deckcraft_components import DeckComponents


def build_pitch_deck():
    output_dir = Path(__file__).resolve().parent / "output"
    output_dir.mkdir(parents=True, exist_ok=True)

    deckcraft_file = output_dir / "deckcraft_fast_pitch.deckcraft"
    pptx_file = output_dir / "deckcraft_fast_pitch.pptx"

    # Select design theme
    theme = THEMES["corporate_navy"]

    with DeckCraftClient(auto_gui=True) as client:
        client.set_outputs(
            save_deckcraft=str(deckcraft_file),
            export_pptx=str(pptx_file)
        )
        comp = DeckComponents(client, theme=theme)

        # ----------------------------------------------------------------------
        # SLIDE 1: Cover / Hero Slide
        # ----------------------------------------------------------------------
        comp.cover_slide(
            title="DeckCraft Fast AI Engine",
            subtitle="Eliminating MCP roundtrip latency for instant presentation generation",
            author="Antigravity Pair Programming Team",
            tag="AI PRODUCTIVITY TOOL",
            date_str="November 2026"
        )

        # ----------------------------------------------------------------------
        # SLIDE 2: Problem vs Solution (Comparison Columns)
        # ----------------------------------------------------------------------
        client.new_slide(layout="blank")
        client.set_background(theme.bg_canvas)
        comp.header(
            title="The Problem vs The Fast Solution",
            subtitle="Why traditional agent MCP loops struggle with rich presentation creation",
            category="PARADIGM SHIFT"
        )
        comp.comparison_columns(
            left_col={
                "tag": "Traditional MCP Loop",
                "title": "50+ Sequential Tool Calls",
                "highlight": False,
                "points": [
                    "1 tool call per shape, text, or color change",
                    "Minutes of waiting per presentation",
                    "Context token bloat & timeout risks",
                    "Hallucinations midway through slide 4"
                ]
            },
            right_col={
                "tag": "DeckCraft Fast Skill",
                "title": "1-Shot Script Execution",
                "highlight": True,
                "points": [
                    "Agent writes 1 complete Python script",
                    "Sub-second execution via native CLI runner",
                    "0 context window inflation & zero latency",
                    "Instant visual feedback in DeckCraft GUI"
                ]
            },
            y=150,
            height=330
        )

        # ----------------------------------------------------------------------
        # SLIDE 3: Key Performance Metrics (KPI Cards)
        # ----------------------------------------------------------------------
        client.new_slide(layout="blank")
        client.set_background(theme.bg_canvas)
        comp.header(
            title="Unmatched Speed & Efficiency",
            subtitle="Quantitative performance improvements across all generation phases",
            category="BENCHMARKS"
        )
        comp.stat_cards_row(
            stats=[
                {
                    "value": "< 1 sec",
                    "label": "Total Generation Time",
                    "subtext": "98% faster than MCP loop",
                    "icon": "zap"
                },
                {
                    "value": "0 Calls",
                    "label": "Intermediate Tool Roundtrips",
                    "subtext": "Zero token overhead",
                    "icon": "chart-bar"
                },
                {
                    "value": "100%",
                    "label": "Native PPTX Compatibility",
                    "subtext": "Ready for PowerPoint & DeckCraft",
                    "icon": "file-text"
                }
            ],
            y=160,
            height=180
        )

        # Add a bonus mini feature row below the stats
        comp.feature_cards_row(
            cards=[
                {
                    "title": "Clean Design Tokens",
                    "body": "Built-in typography and color themes tailored for 16:9 canvas.",
                    "icon": "target"
                },
                {
                    "title": "Reusable Components",
                    "body": "Pre-built headers, cards, grids, and metrics for fast assembly.",
                    "icon": "star"
                }
            ],
            y=360,
            height=130
        )

        # ----------------------------------------------------------------------
        # SLIDE 4: Architecture & Workflow (4 Columns)
        # ----------------------------------------------------------------------
        client.new_slide(layout="blank")
        client.set_background(theme.bg_canvas)
        comp.header(
            title="How the Fast Skill Works",
            subtitle="A clean 4-stage pipeline from user request to interactive presentation",
            category="WORKFLOW"
        )
        comp.feature_cards_row(
            cards=[
                {
                    "badge": "STEP 1",
                    "title": "User Input",
                    "body": "User supplies the topic, outline, or presentation PRD.",
                    "icon": "user"
                },
                {
                    "badge": "STEP 2",
                    "title": "Agent Plan",
                    "body": "Agent structures slides, chooses theme, and plans layout.",
                    "icon": "lightbulb"
                },
                {
                    "badge": "STEP 3",
                    "title": "Python Script",
                    "body": "Agent writes 1 script importing DeckCraftClient & components.",
                    "icon": "file"
                },
                {
                    "badge": "STEP 4",
                    "title": "Instant GUI",
                    "body": "Script runs in 1-shot; DeckCraft GUI pops up with the slides.",
                    "icon": "presentation"
                }
            ],
            y=160,
            height=320
        )

        # ----------------------------------------------------------------------
        # SLIDE 5: Closing CTA
        # ----------------------------------------------------------------------
        comp.closing_slide(
            title="Accelerate Your Presentations",
            subtitle="Build beautiful, executive-ready decks with deckcraft-fast-SKILL.",
            cta_text="Start Building Now",
            contact_info="DeckCraft Community • https://getartcraft.com/apps/deckcraft"
        )

    print(f"\n[Done] Successfully built presentation:")
    print(f" - DeckCraft: {deckcraft_file}")
    print(f" - PPTX:      {pptx_file}")


if __name__ == "__main__":
    build_pitch_deck()
