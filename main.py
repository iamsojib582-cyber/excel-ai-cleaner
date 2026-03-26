"""
╔══════════════════════════════════════════════════════════════╗
║            EXCEL AI CLEANER — by Sajeeb The Analyst         ║
║                        main.py                              ║
║                  Entry point — run this file                ║
╚══════════════════════════════════════════════════════════════╝
"""

import tkinter as tk


# ── THEME COLORS ────────────────────────────────────────────
BG_DARK       = "#0e0b1a"
BG_PANEL      = "#1a1530"
GOLD          = "#c9a84c"
GOLD_BRIGHT   = "#e8c96d"
TEXT_PRIMARY  = "#f0eaff"
TEXT_DIM      = "#7a6f9a"
BORDER        = "#2e2550"


# ════════════════════════════════════════════════════════════
#  SPLASH SCREEN
# ════════════════════════════════════════════════════════════

class SplashScreen:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.overrideredirect(True)        # no title bar
        self.root.configure(bg=BG_DARK)
        self.root.attributes("-topmost", True)

        # ── Center on screen
        sw, sh = root.winfo_screenwidth(), root.winfo_screenheight()
        w, h   = 620, 380
        x, y   = (sw - w) // 2, (sh - h) // 2
        self.root.geometry(f"{w}x{h}+{x}+{y}")

        self._build()
        self._animate_progress()

    # ── BUILD UI ─────────────────────────────────────────────
    def _build(self):
        canvas = tk.Canvas(
            self.root, width=620, height=380,
            bg=BG_DARK, highlightthickness=0)
        canvas.pack(fill="both", expand=True)
        self.canvas = canvas

        # Outer gold border
        canvas.create_rectangle(2, 2, 618, 378, outline=GOLD,   width=2)
        canvas.create_rectangle(6, 6, 614, 374, outline=BORDER, width=1)

        # Logo
        self._draw_logo(canvas, cx=310, cy=95)

        # App title
        canvas.create_text(
            310, 165, text="Excel AI Cleaner",
            fill=GOLD_BRIGHT, font=("Segoe UI", 26, "bold"))

        # Tagline
        canvas.create_text(
            310, 198,
            text="Smart Data Cleaning  •  Powered by Groq AI",
            fill=TEXT_DIM, font=("Segoe UI", 11))

        # Divider
        canvas.create_line(80, 220, 540, 220, fill=BORDER, width=1)

        # Developer label
        canvas.create_text(
            310, 240, text="Developed by",
            fill=TEXT_DIM, font=("Segoe UI", 9))

        canvas.create_text(
            310, 262, text="Sajeeb The Analyst",
            fill=GOLD, font=("Segoe UI", 15, "bold"))

        canvas.create_text(
            310, 285, text="v1.0.0  •  Professional Edition",
            fill=TEXT_DIM, font=("Segoe UI", 8))

        # Progress bar track
        canvas.create_rectangle(
            80, 320, 540, 336,
            fill=BORDER, outline="", width=0)

        # Progress bar fill
        self.progress_bar = canvas.create_rectangle(
            80, 320, 80, 336,
            fill=GOLD, outline="", width=0)

        # Status text
        self.status_text = canvas.create_text(
            310, 352, text="Initializing…",
            fill=TEXT_DIM, font=("Segoe UI", 9))

    # ── LOGO ─────────────────────────────────────────────────
    def _draw_logo(self, canvas: tk.Canvas, cx: int, cy: int):
        # Background circle
        canvas.create_oval(
            cx-45, cy-45, cx+45, cy+45,
            fill=BG_PANEL, outline=GOLD, width=2)

        # Chart bars
        for x1, y1, x2, y2, color in [
            (cx-26, cy+20, cx-14, cy-10, GOLD),
            (cx- 8, cy+20, cx+ 4, cy-24, GOLD_BRIGHT),
            (cx+10, cy+20, cx+22, cy+ 2, GOLD),
        ]:
            canvas.create_rectangle(x1, y1, x2, y2, fill=color, outline="")

        # AI sparkle dots
        for dx, dy, r in [(-32,-28,3),(34,-22,2),(30,18,2),(-30,24,2)]:
            canvas.create_oval(
                cx+dx-r, cy+dy-r, cx+dx+r, cy+dy+r,
                fill=GOLD_BRIGHT, outline="")

        # Sparkle cross lines
        for dx, dy, ldx, ldy in [
            (-32,-28,6,0),(-32,-28,0,6),
            ( 34,-22,5,0),( 34,-22,0,5),
        ]:
            canvas.create_line(
                cx+dx-ldx, cy+dy, cx+dx+ldx, cy+dy,
                fill=GOLD_BRIGHT, width=1)
            canvas.create_line(
                cx+dx, cy+dy-ldy, cx+dx, cy+dy+ldy,
                fill=GOLD_BRIGHT, width=1)

    # ── PROGRESS ANIMATION ───────────────────────────────────
    def _animate_progress(self):
        self._steps = [
            (15,  "Loading UI components…"),
            (35,  "Initializing AI engine…"),
            (55,  "Loading data modules…"),
            (75,  "Setting up workspace…"),
            (95,  "Almost ready…"),
            (100, "Welcome, Sajeeb! 🚀"),
        ]
        self._step_index = 0
        self._run_step()

    def _run_step(self):
        if self._step_index >= len(self._steps):
            self.root.after(600, self._launch)
            return

        pct, msg = self._steps[self._step_index]
        self._step_index += 1

        bar_x = 80 + int((540 - 80) * pct / 100)
        self.canvas.coords(self.progress_bar, 80, 320, bar_x, 336)
        self.canvas.itemconfig(self.status_text, text=msg)

        self.root.after(350 if pct < 100 else 500, self._run_step)

    # ── LAUNCH MAIN APP ──────────────────────────────────────
    def _launch(self):
        self.root.destroy()
        main_root = tk.Tk()
        from ui import ExcelAICleanerApp
        ExcelAICleanerApp(main_root)
        main_root.mainloop()


# ════════════════════════════════════════════════════════════
#  ENTRY POINT
# ════════════════════════════════════════════════════════════

if __name__ == "__main__":
    splash_root = tk.Tk()
    SplashScreen(splash_root)
    splash_root.mainloop()