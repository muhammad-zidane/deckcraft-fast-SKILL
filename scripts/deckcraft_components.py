"""
deckcraft_components.py - High-level reusable presentation components & slide layouts
built for DeckCraft (16:9 canvas: 960x540 pt).
Enables AI agents to assemble executive-ready slides in 1-2 function calls.
"""

from typing import List, Dict, Any, Optional
from design_tokens import (
    CANVAS_WIDTH, CANVAS_HEIGHT,
    SAFE_MARGIN_LEFT, SAFE_MARGIN_RIGHT, SAFE_MARGIN_TOP, SAFE_MARGIN_BOTTOM,
    CONTENT_WIDTH, CONTENT_HEIGHT,
    HEADER_Y, SUBTITLE_Y, BODY_TOP_Y, BODY_HEIGHT,
    FONT_DISPLAY, FONT_COVER_SUBTITLE, FONT_TITLE, FONT_SUBTITLE,
    FONT_CARD_TITLE, FONT_BODY, FONT_BODY_SM, FONT_CAPTION,
    FONT_KPI_NUMBER, FONT_KPI_LABEL,
    ThemeColors, THEMES, calculate_columns
)


class DeckComponents:
    def __init__(self, client: Any, theme: Optional[ThemeColors] = None):
        self.client = client
        self.theme = theme or THEMES["corporate_navy"]

    # ==========================================================================
    # 1. Slide Headers & Dividers
    # ==========================================================================
    def header(
        self,
        title: str,
        subtitle: Optional[str] = None,
        category: Optional[str] = None
    ):
        """Standard top-aligned slide header section with optional category pill and subtitle."""
        current_y = HEADER_Y

        # Category Pill / Tag
        if category:
            self.client.add_badge(
                text=category.upper(),
                x=SAFE_MARGIN_LEFT,
                y=current_y,
                width=max(90, len(category) * 8 + 20),
                height=22,
                bg_color=self.theme.accent_surface,
                text_color=self.theme.accent_primary,
                font_size=10.0,
                font_family=self.theme.font_heading
            )
            current_y += 28

        # Main Slide Title
        self.client.add_textbox(
            rect=[SAFE_MARGIN_LEFT, current_y, CONTENT_WIDTH, 36],
            text=title,
            font_size=FONT_TITLE,
            font_color=self.theme.text_primary,
            font_family=self.theme.font_heading,
            bold=True,
            align="left"
        )
        current_y += 38

        # Optional Subtitle
        if subtitle:
            self.client.add_textbox(
                rect=[SAFE_MARGIN_LEFT, current_y, CONTENT_WIDTH, 24],
                text=subtitle,
                font_size=FONT_SUBTITLE,
                font_color=self.theme.text_secondary,
                font_family=self.theme.font_body,
                bold=False,
                align="left"
            )

    # ==========================================================================
    # 2. Cover / Hero Slide Layout
    # ==========================================================================
    def cover_slide(
        self,
        title: str,
        subtitle: str,
        author: Optional[str] = None,
        tag: str = "PRESENTATION",
        date_str: Optional[str] = None
    ):
        """Creates a modern, punchy title / hero cover slide."""
        self.client.new_slide(layout="blank")
        self.client.set_background(self.theme.bg_canvas)

        # Tag Badge
        self.client.add_badge(
            text=tag.upper(),
            x=SAFE_MARGIN_LEFT,
            y=130,
            width=max(110, len(tag) * 8 + 24),
            height=26,
            bg_color=self.theme.accent_surface,
            text_color=self.theme.accent_primary,
            font_size=11.0,
            font_family=self.theme.font_heading
        )

        # Big Display Title
        self.client.add_textbox(
            rect=[SAFE_MARGIN_LEFT, 170, 780, 100],
            text=title,
            font_size=FONT_DISPLAY,
            font_color=self.theme.text_primary,
            font_family=self.theme.font_heading,
            bold=True,
            align="left"
        )

        # Subtitle
        self.client.add_textbox(
            rect=[SAFE_MARGIN_LEFT, 280, 740, 50],
            text=subtitle,
            font_size=FONT_COVER_SUBTITLE,
            font_color=self.theme.text_secondary,
            font_family=self.theme.font_body,
            bold=False,
            align="left"
        )

        # Subtle Accent Separator Line
        self.client.add_shape(
            preset="rect",
            rect=[SAFE_MARGIN_LEFT, 350, 80, 4],
            fill=self.theme.accent_primary,
            stroke=None
        )

        # Author / Metadata Footer
        meta_parts = []
        if author:
            meta_parts.append(author)
        if date_str:
            meta_parts.append(date_str)
        if meta_parts:
            self.client.add_textbox(
                rect=[SAFE_MARGIN_LEFT, 440, CONTENT_WIDTH, 24],
                text="  •  ".join(meta_parts),
                font_size=FONT_BODY_SM,
                font_color=self.theme.text_secondary,
                font_family=self.theme.font_body,
                bold=True,
                align="left"
            )

    # ==========================================================================
    # 3. KPI / Stat Metrics Row
    # ==========================================================================
    def stat_cards_row(
        self,
        stats: List[Dict[str, Any]],
        y: float = 160.0,
        height: float = 170.0
    ):
        """
        Creates a row of statistical KPI cards.
        Each stat dict: {"value": "$24M", "label": "Annual Run Rate", "subtext": "+120% YoY", "icon": "trending-up"}
        """
        cols = calculate_columns(len(stats), start_x=SAFE_MARGIN_LEFT, total_width=CONTENT_WIDTH, gap=20.0)

        for col, stat in zip(cols, stats):
            cx = col["x"]
            cw = col["width"]

            # Card Container
            self.client.add_card(
                rect=[cx, y, cw, height],
                fill=self.theme.bg_card,
                stroke=self.theme.border_card,
                stroke_width=1.0
            )

            # Optional Icon
            icon_name = stat.get("icon")
            content_x = cx + 24
            if icon_name:
                self.client.add_icon(
                    name=icon_name,
                    rect=[cx + cw - 44, y + 20, 24, 24],
                    fill=self.theme.accent_primary
                )

            # Metric Block
            self.client.add_stat_metric(
                value=stat.get("value", "0"),
                label=stat.get("label", ""),
                rect=[content_x, y + 24, cw - 48, height - 48],
                value_color=self.theme.text_primary,
                label_color=self.theme.text_secondary,
                value_size=36.0,
                label_size=13.0,
                font_family=self.theme.font_heading,
                subtext=stat.get("subtext"),
                subtext_color=self.theme.accent_success
            )

    # ==========================================================================
    # 4. Feature / Content Cards Grid
    # ==========================================================================
    def feature_cards_row(
        self,
        cards: List[Dict[str, Any]],
        y: float = 150.0,
        height: float = 320.0
    ):
        """
        Evenly spaced columns of feature/benefit cards.
        Each card dict: {"title": "Fast Execution", "body": "...", "badge": "Core", "icon": "zap"}
        """
        cols = calculate_columns(len(cards), start_x=SAFE_MARGIN_LEFT, total_width=CONTENT_WIDTH, gap=20.0)

        for col, card in zip(cols, cards):
            cx = col["x"]
            cw = col["width"]

            # Card background
            self.client.add_card(
                rect=[cx, y, cw, height],
                fill=self.theme.bg_card,
                stroke=self.theme.border_card,
                stroke_width=1.0
            )

            inner_x = cx + 24
            inner_w = cw - 48
            curr_y = y + 24

            # Optional Badge
            if "badge" in card:
                self.client.add_badge(
                    text=card["badge"],
                    x=inner_x,
                    y=curr_y,
                    width=max(70, len(card["badge"]) * 7 + 16),
                    height=20,
                    bg_color=self.theme.accent_surface,
                    text_color=self.theme.accent_primary,
                    font_size=10.0,
                    font_family=self.theme.font_heading
                )
                curr_y += 30

            # Optional Icon
            if "icon" in card:
                self.client.add_icon(
                    name=card["icon"],
                    rect=[inner_x, curr_y, 28, 28],
                    fill=self.theme.accent_primary
                )
                curr_y += 38

            # Card Title
            self.client.add_textbox(
                rect=[inner_x, curr_y, inner_w, 28],
                text=card.get("title", ""),
                font_size=FONT_CARD_TITLE,
                font_color=self.theme.text_primary,
                font_family=self.theme.font_heading,
                bold=True,
                align="left"
            )
            curr_y += 34

            # Card Body
            self.client.add_textbox(
                rect=[inner_x, curr_y, inner_w, height - (curr_y - y) - 20],
                text=card.get("body", ""),
                font_size=FONT_BODY,
                font_color=self.theme.text_secondary,
                font_family=self.theme.font_body,
                bold=False,
                align="left"
            )

    # ==========================================================================
    # 5. Side-by-Side Comparison Columns
    # ==========================================================================
    def comparison_columns(
        self,
        left_col: Dict[str, Any],
        right_col: Dict[str, Any],
        y: float = 150.0,
        height: float = 320.0
    ):
        """
        Two prominent comparison columns (e.g. 'Old Way vs New Way' or 'Problem vs Solution').
        Left/Right dict: {"tag": "BEFORE", "title": "...", "points": ["...", "..."], "highlight": False}
        """
        cols = calculate_columns(2, start_x=SAFE_MARGIN_LEFT, total_width=CONTENT_WIDTH, gap=24.0)

        for col, col_data in zip(cols, [left_col, right_col]):
            cx = col["x"]
            cw = col["width"]
            is_highlight = col_data.get("highlight", False)

            card_bg = self.theme.accent_surface if is_highlight else self.theme.bg_card
            card_border = self.theme.accent_primary if is_highlight else self.theme.border_card

            self.client.add_card(
                rect=[cx, y, cw, height],
                fill=card_bg,
                stroke=card_border,
                stroke_width=1.5 if is_highlight else 1.0
            )

            inner_x = cx + 28
            inner_w = cw - 56
            curr_y = y + 24

            # Tag
            tag_text = col_data.get("tag", "CATEGORY")
            self.client.add_badge(
                text=tag_text.upper(),
                x=inner_x,
                y=curr_y,
                width=max(80, len(tag_text) * 7 + 16),
                height=22,
                bg_color=self.theme.accent_primary if is_highlight else self.theme.bg_card_alt,
                text_color="#FFFFFF" if is_highlight else self.theme.text_secondary,
                font_size=10.0,
                font_family=self.theme.font_heading
            )
            curr_y += 32

            # Title
            self.client.add_textbox(
                rect=[inner_x, curr_y, inner_w, 28],
                text=col_data.get("title", ""),
                font_size=FONT_CARD_TITLE,
                font_color=self.theme.text_primary,
                font_family=self.theme.font_heading,
                bold=True,
                align="left"
            )
            curr_y += 36

            # Bullet points
            points = col_data.get("points", [])
            for pt in points:
                # Bullet icon / symbol
                bullet_icon = "check" if is_highlight else "x"
                bullet_color = self.theme.accent_success if is_highlight else self.theme.accent_warning
                self.client.add_icon(
                    name=bullet_icon,
                    rect=[inner_x, curr_y + 2, 16, 16],
                    fill=bullet_color
                )
                # Point text
                self.client.add_textbox(
                    rect=[inner_x + 24, curr_y, inner_w - 24, 24],
                    text=pt,
                    font_size=FONT_BODY,
                    font_color=self.theme.text_secondary,
                    font_family=self.theme.font_body,
                    bold=False,
                    align="left"
                )
                curr_y += 32

    # ==========================================================================
    # 6. Horizontal Timeline / Roadmap
    # ==========================================================================
    def timeline_steps(
        self,
        steps: List[Dict[str, Any]],
        y: float = 200.0,
        height: float = 240.0
    ):
        """
        A chronological roadmap / milestone progression.
        Each step: {"phase": "Q1 2026", "title": "Alpha Release", "desc": "Initial core engine launch"}
        """
        cols = calculate_columns(len(steps), start_x=SAFE_MARGIN_LEFT, total_width=CONTENT_WIDTH, gap=20.0)

        # Draw connecting line
        line_y = y + 20
        self.client.add_shape(
            preset="rect",
            rect=[SAFE_MARGIN_LEFT + 20, line_y, CONTENT_WIDTH - 40, 3],
            fill=self.theme.border_card,
            stroke=None
        )

        for idx, (col, step) in enumerate(zip(cols, steps)):
            cx = col["x"]
            cw = col["width"]

            # Step Marker Circle
            self.client.add_shape(
                preset="ellipse",
                rect=[cx + 12, line_y - 8, 20, 20],
                fill=self.theme.accent_primary,
                stroke="#FFFFFF",
                stroke_width=2.0
            )

            # Step Container Card
            card_y = y + 40
            self.client.add_card(
                rect=[cx, card_y, cw, height - 40],
                fill=self.theme.bg_card,
                stroke=self.theme.border_card,
                stroke_width=1.0
            )

            inner_x = cx + 16
            inner_w = cw - 32

            # Phase label
            self.client.add_textbox(
                rect=[inner_x, card_y + 16, inner_w, 20],
                text=step.get("phase", f"STEP {idx+1}"),
                font_size=FONT_CAPTION,
                font_color=self.theme.accent_primary,
                font_family=self.theme.font_heading,
                bold=True,
                align="left"
            )

            # Title
            self.client.add_textbox(
                rect=[inner_x, card_y + 38, inner_w, 24],
                text=step.get("title", ""),
                font_size=FONT_BODY,
                font_color=self.theme.text_primary,
                font_family=self.theme.font_heading,
                bold=True,
                align="left"
            )

            # Description
            self.client.add_textbox(
                rect=[inner_x, card_y + 66, inner_w, 90],
                text=step.get("desc", ""),
                font_size=FONT_BODY_SM,
                font_color=self.theme.text_secondary,
                font_family=self.theme.font_body,
                bold=False,
                align="left"
            )

    # ==========================================================================
    # 7. Closing / CTA Slide
    # ==========================================================================
    def closing_slide(
        self,
        title: str = "Thank You",
        subtitle: Optional[str] = None,
        cta_text: Optional[str] = "Get Started",
        contact_info: Optional[str] = None
    ):
        """Creates a high-impact closing / call-to-action slide."""
        self.client.new_slide(layout="blank")
        self.client.set_background(self.theme.bg_canvas)

        center_x = CANVAS_WIDTH / 2
        
        # Center Title
        self.client.add_textbox(
            rect=[SAFE_MARGIN_LEFT, 160, CONTENT_WIDTH, 60],
            text=title,
            font_size=FONT_DISPLAY,
            font_color=self.theme.text_primary,
            font_family=self.theme.font_heading,
            bold=True,
            align="center"
        )

        # Center Subtitle
        if subtitle:
            self.client.add_textbox(
                rect=[SAFE_MARGIN_LEFT, 230, CONTENT_WIDTH, 40],
                text=subtitle,
                font_size=FONT_COVER_SUBTITLE,
                font_color=self.theme.text_secondary,
                font_family=self.theme.font_body,
                bold=False,
                align="center"
            )

        # CTA Button
        if cta_text:
            btn_w = 200
            btn_h = 44
            btn_x = center_x - (btn_w / 2)
            btn_y = 290
            self.client.add_shape(
                preset="roundRect",
                rect=[btn_x, btn_y, btn_w, btn_h],
                text=cta_text,
                fill=self.theme.accent_primary,
                stroke=None,
                text_color="#FFFFFF",
                font_size=FONT_BODY,
                font_family=self.theme.font_heading,
                bold=True,
                align="center"
            )

        # Contact Info Footer
        if contact_info:
            self.client.add_textbox(
                rect=[SAFE_MARGIN_LEFT, 380, CONTENT_WIDTH, 24],
                text=contact_info,
                font_size=FONT_BODY_SM,
                font_color=self.theme.text_secondary,
                font_family=self.theme.font_body,
                bold=False,
                align="center"
            )
