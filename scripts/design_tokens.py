"""
design_tokens.py - Design system, color palettes, typography scales, and grid dimensions
for DeckCraft presentation slides (16:9 canvas: 960x540 pt).
"""

from dataclasses import dataclass
from typing import Dict, Any, List, Optional


# ==============================================================================
# 1. CANVAS & SAFE MARGIN CONSTANTS (DeckCraft Standard 16:9 = 960x540 pt)
# ==============================================================================
CANVAS_WIDTH = 960
CANVAS_HEIGHT = 540

SAFE_MARGIN_LEFT = 60
SAFE_MARGIN_RIGHT = 60
SAFE_MARGIN_TOP = 48
SAFE_MARGIN_BOTTOM = 48

CONTENT_WIDTH = CANVAS_WIDTH - (SAFE_MARGIN_LEFT + SAFE_MARGIN_RIGHT)  # 840 pt
CONTENT_HEIGHT = CANVAS_HEIGHT - (SAFE_MARGIN_TOP + SAFE_MARGIN_BOTTOM)  # 444 pt

# Standard Y Positions
HEADER_Y = 48
SUBTITLE_Y = 96
BODY_TOP_Y = 150
BODY_HEIGHT = 330
FOOTER_Y = 500


# ==============================================================================
# 2. TYPOGRAPHY SCALES
# ==============================================================================
FONT_DISPLAY = 42       # Big hero title on cover slide
FONT_COVER_SUBTITLE = 20 # Subtitle on cover slide
FONT_TITLE = 28         # Standard slide header
FONT_SUBTITLE = 16      # Sub-header under slide title
FONT_CARD_TITLE = 18    # Card / Box header
FONT_BODY = 14          # Regular descriptive text
FONT_BODY_SM = 12       # Secondary text, list sub-items
FONT_CAPTION = 10       # Footnotes, sources, tags
FONT_KPI_NUMBER = 44    # Big statistical numbers
FONT_KPI_LABEL = 13     # Label under big numbers


# ==============================================================================
# 3. COLOR THEMES
# ==============================================================================
@dataclass
class ThemeColors:
    name: str
    bg_canvas: str         # Slide background color
    bg_card: str           # Card/Container background color
    bg_card_alt: str       # Secondary container / badge background
    border_card: str       # Border/stroke color for containers
    text_primary: str      # Main title & body text
    text_secondary: str    # Subtitles & muted labels
    text_accent: str       # Highlighted text
    accent_primary: str    # Main brand color (buttons, hero tags, highlights)
    accent_secondary: str  # Secondary accent (pills, secondary metrics)
    accent_surface: str    # Tinted pill / badge background
    accent_warning: str    # Alert/highlight color
    accent_success: str    # Growth/positive metric color
    font_heading: str = "Inter"
    font_body: str = "Inter"


# Pre-defined professional color themes
THEMES: Dict[str, ThemeColors] = {
    "corporate_navy": ThemeColors(
        name="Corporate Navy",
        bg_canvas="#F8FAFC",
        bg_card="#FFFFFF",
        bg_card_alt="#F1F5F9",
        border_card="#E2E8F0",
        text_primary="#0F172A",
        text_secondary="#475569",
        text_accent="#2563EB",
        accent_primary="#1E40AF",
        accent_secondary="#3B82F6",
        accent_surface="#EFF6FF",
        accent_warning="#F59E0B",
        accent_success="#10B981",
        font_heading="Inter",
        font_body="Inter",
    ),
    "modern_dark": ThemeColors(
        name="Modern Dark",
        bg_canvas="#0B0F19",
        bg_card="#151C2C",
        bg_card_alt="#1E293B",
        border_card="#2A354D",
        text_primary="#F8FAFC",
        text_secondary="#94A3B8",
        text_accent="#38BDF8",
        accent_primary="#0284C7",
        accent_secondary="#38BDF8",
        accent_surface="#0C4A6E",
        accent_warning="#FBBF24",
        accent_success="#34D399",
        font_heading="Inter",
        font_body="Inter",
    ),
    "clean_slate": ThemeColors(
        name="Clean Slate",
        bg_canvas="#FFFFFF",
        bg_card="#F9FAFB",
        bg_card_alt="#F3F4F6",
        border_card="#E5E7EB",
        text_primary="#111827",
        text_secondary="#4B5563",
        text_accent="#4F46E5",
        accent_primary="#4F46E5",
        accent_secondary="#818CF8",
        accent_surface="#EEF2FF",
        accent_warning="#F59E0B",
        accent_success="#10B981",
        font_heading="Inter",
        font_body="Inter",
    ),
    "vibrant_purple": ThemeColors(
        name="Vibrant Purple",
        bg_canvas="#0E0A1E",
        bg_card="#1B1438",
        bg_card_alt="#271E4E",
        border_card="#3A2D73",
        text_primary="#FFFFFF",
        text_secondary="#A594FD",
        text_accent="#F43F5E",
        accent_primary="#8B5CF6",
        accent_secondary="#A78BFA",
        accent_surface="#382574",
        accent_warning="#FBBF24",
        accent_success="#10B981",
        font_heading="Montserrat",
        font_body="Lato",
    ),
    "emerald_growth": ThemeColors(
        name="Emerald Growth",
        bg_canvas="#F0FDF4",
        bg_card="#FFFFFF",
        bg_card_alt="#DCFCE7",
        border_card="#BBF7D0",
        text_primary="#064E3B",
        text_secondary="#166534",
        text_accent="#059669",
        accent_primary="#059669",
        accent_secondary="#10B981",
        accent_surface="#ECFDF5",
        accent_warning="#F59E0B",
        accent_success="#047857",
        font_heading="Inter",
        font_body="Inter",
    ),
    "warm_editorial": ThemeColors(
        name="Warm Editorial",
        bg_canvas="#FFFDF7",
        bg_card="#FFFFFF",
        bg_card_alt="#FEF3C7",
        border_card="#E7E5E4",
        text_primary="#292524",
        text_secondary="#78716C",
        text_accent="#B45309",
        accent_primary="#B45309",
        accent_secondary="#D97706",
        accent_surface="#FEF3C7",
        accent_warning="#D97706",
        accent_success="#059669",
        font_heading="Playfair Display",
        font_body="Lato",
    ),
}


# ==============================================================================
# 4. GRID LAYOUT CALCULATION UTILITIES
# ==============================================================================
def calculate_columns(
    total_cols: int,
    start_x: float = SAFE_MARGIN_LEFT,
    total_width: float = CONTENT_WIDTH,
    gap: float = 20.0
) -> List[Dict[str, float]]:
    """
    Computes horizontal coordinates [x, width] for evenly spaced columns.
    
    Example:
        cols = calculate_columns(3)
        # Returns [{'index': 0, 'x': 60, 'width': 266.6}, ...]
    """
    if total_cols <= 0:
        return []
    total_gaps = gap * (total_cols - 1)
    col_width = (total_width - total_gaps) / total_cols
    
    result = []
    current_x = start_x
    for i in range(total_cols):
        result.append({
            "index": i,
            "x": round(current_x, 1),
            "width": round(col_width, 1)
        })
        current_x += col_width + gap
    return result
