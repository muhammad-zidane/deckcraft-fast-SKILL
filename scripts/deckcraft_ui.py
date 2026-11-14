"""
deckcraft_ui.py - Production Fast Automation Library for DeckCraft.
Supports Hybrid execution:
  1. Live TCP Streaming (when DeckCraft GUI is running with `--control <port>`).
  2. Sub-second Batch Execution via `deckcraft-cli run --cmd-file` + automatic GUI launch.
Exports native `.deckcraft` and Microsoft PowerPoint `.pptx` presentations.
"""

import json
import os
import socket
import subprocess
import tempfile
import time
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple, Union


DEFAULT_DECKCRAFT_CLI = r"C:\Program Files\DeckCraft\deckcraft-cli.exe"
DEFAULT_DECKCRAFT_APP = r"C:\Program Files\DeckCraft\deckcraft.exe"


class DeckCraftError(RuntimeError):
    """Raised when DeckCraft CLI or RPC command execution fails."""
    pass


class DeckCraftClient:
    def __init__(
        self,
        host: str = "127.0.0.1",
        port: int = 7980,
        cli_path: Optional[str] = None,
        app_path: Optional[str] = None,
        timeout: float = 10.0,
        auto_gui: bool = True
    ):
        self.host = host
        self.port = port
        self.cli_path = cli_path or DEFAULT_DECKCRAFT_CLI
        self.app_path = app_path or DEFAULT_DECKCRAFT_APP
        self.timeout = timeout
        self.auto_gui = auto_gui

        self.sock: Optional[socket.socket] = None
        self.is_live: bool = False
        self._req_id = 0
        self._cmd_queue: List[Tuple[str, Dict[str, Any]]] = []
        self._target_save_path: Optional[str] = None
        self._target_pptx_path: Optional[str] = None
        self._target_png_path: Optional[str] = None
        self._executed_commands_count = 0

        # Probe for live control port
        self._probe_live_connection()

    def _probe_live_connection(self):
        """Attempts to connect to DeckCraft GUI control port."""
        s = None
        try:
            s = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
            s.settimeout(0.5)
            s.connect((self.host, self.port))
            # Test ping
            req = {
                "jsonrpc": "2.0",
                "id": 1,
                "method": "engine.execute",
                "params": {"command": "file.info", "params": {}}
            }
            s.sendall((json.dumps(req) + "\n").encode("utf-8"))
            buf = s.recv(1024)
            if buf:
                self.sock = s
                self.is_live = True
                print(f"[DeckCraft] Connected to live GUI on {self.host}:{self.port}")
                return
        except Exception:
            if s:
                try:
                    s.close()
                except Exception:
                    pass
            self.sock = None
            self.is_live = False

    def add_command(self, command: str, params: Optional[Dict[str, Any]] = None):
        """Dispatches command immediately if live streaming, or enqueues for 1-shot batch execution."""
        if params is None:
            params = {}
        self._cmd_queue.append((command, params))

        if self.is_live and self.sock:
            self._send_rpc(command, params)

    def _send_rpc(self, command: str, params: Dict[str, Any]) -> Any:
        """Sends command over live TCP JSON-RPC."""
        self._req_id += 1
        req = {
            "jsonrpc": "2.0",
            "id": self._req_id,
            "method": "engine.execute",
            "params": {
                "command": command,
                "params": params
            }
        }
        payload = (json.dumps(req) + "\n").encode("utf-8")
        try:
            self.sock.sendall(payload)
            buf = b""
            while True:
                chunk = self.sock.recv(4096)
                if not chunk:
                    break
                buf += chunk
                if b"\n" in buf:
                    break
            line = buf.decode("utf-8", errors="ignore").strip()
            if line:
                data = json.loads(line)
                return data.get("result")
        except Exception as e:
            print(f"[DeckCraft Warning] Live RPC failed for '{command}': {e}. Switching to batch mode.")
            self.is_live = False
            if self.sock:
                try:
                    self.sock.close()
                except Exception:
                    pass
                self.sock = None
        return None

    # ==========================================================================
    # Presentation & Slide Management
    # ==========================================================================
    def new_presentation(
        self,
        theme: str = "Slate",
        size: Optional[List[int]] = None,
        blank: bool = True
    ):
        """Creates a new presentation canvas."""
        params: Dict[str, Any] = {"theme": theme, "blank": blank}
        if size:
            params["size"] = size
        self.add_command("file.new", params)

    def new_slide(
        self,
        layout: str = "blank",
        title: Optional[str] = None,
        body: Optional[str] = None,
        at: Optional[int] = None
    ):
        """Creates a new slide."""
        params: Dict[str, Any] = {"layout": layout}
        if title is not None:
            params["title"] = title
        if body is not None:
            params["body"] = body
        if at is not None:
            params["at"] = at
        self.add_command("slide.new", params)

    def set_background(self, color: str, index: Optional[int] = None):
        """Sets background color of the current or specified slide."""
        params: Dict[str, Any] = {"color": color}
        if index is not None:
            params["index"] = index
        self.add_command("design.background", params)

    def set_theme(self, theme: str):
        """Applies a theme."""
        self.add_command("design.theme", {"theme": theme})

    # ==========================================================================
    # Shapes, Containers & Cards
    # ==========================================================================
    def add_shape(
        self,
        preset: str = "roundRect",
        rect: Optional[List[float]] = None,
        text: Optional[str] = None,
        fill: Optional[str] = None,
        stroke: Optional[str] = None,
        stroke_width: float = 1.0,
        text_color: Optional[str] = None,
        font_size: Optional[float] = None,
        font_family: Optional[str] = None,
        bold: bool = False,
        align: Optional[str] = None
    ):
        """Inserts a shape and applies optional fills, strokes, and text styles."""
        if rect is None:
            rect = [60, 60, 200, 100]
        params: Dict[str, Any] = {"preset": preset, "rect": rect}
        if text:
            params["text"] = text
        self.add_command("shape.insert", params)

        if fill:
            self.add_command("shape.fill", {"color": fill})
        if stroke:
            self.add_command("shape.line", {"color": stroke, "width": stroke_width})

        # Format text inside shape
        if font_family:
            self.add_command("format.font", {"family": font_family})
        if font_size:
            self.add_command("format.size", {"size": font_size})
        if text_color:
            self.add_command("format.color", {"color": text_color})
        if bold:
            self.add_command("format.bold", {"bold": True})
        if align:
            self.add_command("format.align", {"align": align})

    def add_card(
        self,
        rect: List[float],
        fill: str = "#FFFFFF",
        stroke: Optional[str] = "#E2E8F0",
        stroke_width: float = 1.0,
        preset: str = "roundRect"
    ):
        """Inserts a clean container card for structured slide content."""
        self.add_shape(
            preset=preset,
            rect=rect,
            fill=fill,
            stroke=stroke,
            stroke_width=stroke_width
        )

    # ==========================================================================
    # Text, Typography & Badges
    # ==========================================================================
    def add_textbox(
        self,
        rect: List[float],
        text: str,
        font_size: float = 16.0,
        font_color: str = "#0F172A",
        font_family: str = "Inter",
        bold: bool = False,
        italic: bool = False,
        align: str = "left"
    ):
        """Inserts a text box with full typographic formatting."""
        self.add_command("insert.textBox", {"rect": rect, "text": text})
        self.add_command("format.font", {"family": font_family})
        self.add_command("format.size", {"size": font_size})
        self.add_command("format.color", {"color": font_color})
        if bold:
            self.add_command("format.bold", {"bold": True})
        if italic:
            self.add_command("format.italic", {"italic": True})
        if align:
            self.add_command("format.align", {"align": align})

    def add_badge(
        self,
        text: str,
        x: float,
        y: float,
        width: float = 100.0,
        height: float = 24.0,
        bg_color: str = "#EFF6FF",
        text_color: str = "#1E40AF",
        font_size: float = 11.0,
        font_family: str = "Inter",
        preset: str = "roundRect"
    ):
        """Inserts a pill/badge component for categories, tags, or status."""
        self.add_shape(
            preset=preset,
            rect=[x, y, width, height],
            text=text,
            fill=bg_color,
            stroke=None,
            text_color=text_color,
            font_size=font_size,
            font_family=font_family,
            bold=True,
            align="center"
        )

    # ==========================================================================
    # High-Level Visual Components (Icons, Metrics, Tables, Charts)
    # ==========================================================================
    def add_icon(
        self,
        name: str,
        rect: List[float],
        fill: Optional[str] = None
    ):
        """Inserts an icon from the built-in library (e.g. 'star', 'chart-bar', 'zap', 'shield')."""
        self.add_command("insert.icon", {"name": name, "rect": rect})
        if fill:
            self.add_command("shape.fill", {"color": fill})

    def add_stat_metric(
        self,
        value: str,
        label: str,
        rect: List[float],
        value_color: str = "#0F172A",
        label_color: str = "#64748B",
        value_size: float = 38.0,
        label_size: float = 13.0,
        font_family: str = "Inter",
        subtext: Optional[str] = None,
        subtext_color: str = "#10B981"
    ):
        """Inserts a KPI stat metric block (big number + descriptive label)."""
        x, y, w, h = rect
        # Number
        num_h = value_size + 8
        self.add_textbox(
            rect=[x, y, w, num_h],
            text=value,
            font_size=value_size,
            font_color=value_color,
            font_family=font_family,
            bold=True,
            align="left"
        )
        # Label
        lbl_y = y + num_h + 2
        lbl_h = label_size + 6
        self.add_textbox(
            rect=[x, lbl_y, w, lbl_h],
            text=label,
            font_size=label_size,
            font_color=label_color,
            font_family=font_family,
            bold=False,
            align="left"
        )
        # Optional Subtext / Badge
        if subtext:
            sub_y = lbl_y + lbl_h + 4
            self.add_textbox(
                rect=[x, sub_y, w, 20],
                text=subtext,
                font_size=11.0,
                font_color=subtext_color,
                font_family=font_family,
                bold=True,
                align="left"
            )

    def add_chart(
        self,
        chart_type: str = "column",
        rect: Optional[List[float]] = None,
        categories: Optional[List[str]] = None,
        series: Optional[List[Dict[str, Any]]] = None
    ):
        """
        Inserts a native presentation chart.
        Types: 'column', 'bar', 'line', 'pie', 'doughnut', 'area'
        """
        if rect is None:
            rect = [60, 160, 420, 260]
        params: Dict[str, Any] = {"type": chart_type, "rect": rect}
        if categories:
            params["categories"] = categories
        if series:
            params["series"] = series
        self.add_command("insert.chart", params)

    def add_table(
        self,
        rect: List[float],
        rows: int,
        cols: int,
        headers: Optional[List[str]] = None,
        data: Optional[List[List[str]]] = None,
        style: Optional[str] = None
    ):
        """Inserts and populates a styled presentation table."""
        self.add_command("insert.table", {"rect": rect, "rows": rows, "cols": cols})
        if style:
            self.add_command("table.style", {"style": style})

        # Fill table cell content if provided
        if headers:
            for c_idx, h_text in enumerate(headers[:cols]):
                self.add_command("text.set", {"text": str(h_text), "cell": [0, c_idx]})

        if data:
            start_row = 1 if headers else 0
            for r_idx, row in enumerate(data):
                actual_row = start_row + r_idx
                if actual_row >= rows:
                    break
                for c_idx, val in enumerate(row[:cols]):
                    self.add_command("text.set", {"text": str(val), "cell": [actual_row, c_idx]})

    # ==========================================================================
    # Output Targets & Flushing
    # ==========================================================================
    def set_outputs(
        self,
        save_deckcraft: Optional[str] = None,
        export_pptx: Optional[str] = None,
        export_png: Optional[str] = None
    ):
        """Sets default target output paths."""
        if save_deckcraft:
            self._target_save_path = str(Path(save_deckcraft).resolve())
        if export_pptx:
            self._target_pptx_path = str(Path(export_pptx).resolve())
        if export_png:
            self._target_png_path = str(Path(export_png).resolve())

    def flush(
        self,
        save_deckcraft: Optional[str] = None,
        export_pptx: Optional[str] = None,
        export_png: Optional[str] = None
    ) -> Dict[str, Any]:
        """
        Executes all queued commands in sub-second batch mode via `deckcraft-cli run --cmd-file`.
        Saves native `.deckcraft`, exports `.pptx`, and opens GUI if requested.
        """
        save_path = save_deckcraft or self._target_save_path
        pptx_path = export_pptx or self._target_pptx_path
        png_path = export_png or self._target_png_path

        if not self._cmd_queue:
            return {"status": "empty", "commands": 0}

        # If already streaming live, just save/export
        if self.is_live:
            if save_path:
                self._send_rpc("file.saveAs", {"path": save_path})
            if pptx_path:
                self._send_rpc("file.export", {"path": pptx_path, "format": "pptx"})
            if png_path:
                self._send_rpc("file.export", {"path": png_path, "format": "png"})
            return {"status": "ok", "mode": "live", "commands": len(self._cmd_queue)}

        # Batch execution via deckcraft-cli
        with tempfile.NamedTemporaryFile("w", encoding="utf-8", suffix=".txt", delete=False) as f:
            cmd_file_path = f.name
            for cmd_name, cmd_params in self._cmd_queue:
                if cmd_params:
                    line = f"{cmd_name}={json.dumps(cmd_params, ensure_ascii=False)}\n"
                else:
                    line = f"{cmd_name}\n"
                f.write(line)

        try:
            cli_args = [self.cli_path, "run", "--cmd-file", cmd_file_path]
            if save_path:
                Path(save_path).parent.mkdir(parents=True, exist_ok=True)
                cli_args.extend(["--save", save_path])
            if pptx_path:
                Path(pptx_path).parent.mkdir(parents=True, exist_ok=True)
                cli_args.extend(["--export", pptx_path])
            elif png_path:
                Path(png_path).parent.mkdir(parents=True, exist_ok=True)
                cli_args.extend(["--export", png_path])

            t0 = time.perf_counter()
            result = subprocess.run(
                cli_args,
                capture_output=True,
                text=True,
                timeout=self.timeout
            )
            elapsed = time.perf_counter() - t0

            if result.returncode != 0:
                raise DeckCraftError(
                    f"deckcraft-cli run failed (exit {result.returncode}):\n{result.stderr or result.stdout}"
                )

            # If PPTX was requested and we also have save_path, export PPTX in second quick pass if not already done
            if pptx_path and save_path and not Path(pptx_path).exists():
                subprocess.run(
                    [self.cli_path, "convert", save_path, pptx_path],
                    capture_output=True,
                    timeout=self.timeout
                )

            # Automatically launch DeckCraft GUI to display presentation
            if self.auto_gui and (save_path or pptx_path):
                display_target = save_path if save_path and Path(save_path).exists() else pptx_path
                self.show_in_gui(display_target)

            return {
                "status": "ok",
                "mode": "batch",
                "commands": len(self._cmd_queue),
                "elapsed_sec": round(elapsed, 3),
                "saved_deckcraft": save_path,
                "exported_pptx": pptx_path,
                "exported_png": png_path
            }

        finally:
            try:
                os.remove(cmd_file_path)
            except Exception:
                pass

    def show_in_gui(self, file_path: Optional[str] = None):
        """Brings up or launches DeckCraft GUI displaying the target presentation."""
        target = file_path or self._target_save_path or self._target_pptx_path
        if not target or not Path(target).exists():
            return

        try:
            # Check if deckcraft is already running
            chk = subprocess.run(
                ["powershell", "-NoProfile", "-Command", "Get-Process deckcraft -ErrorAction SilentlyContinue"],
                capture_output=True,
                text=True
            )
            # If not running or target specified, launch DeckCraft GUI with target presentation
            subprocess.Popen(
                [self.app_path, str(Path(target).resolve())],
                shell=False
            )
            print(f"[DeckCraft] Presentation loaded in DeckCraft GUI: {target}")
        except Exception as e:
            print(f"[DeckCraft Warning] Could not launch GUI: {e}")

    def __enter__(self):
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        if exc_type is None:
            self.flush()
        if self.sock:
            try:
                self.sock.close()
            except Exception:
                pass
            self.sock = None
        return False
