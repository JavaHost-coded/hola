"""Small Tkinter visual helpers used by the Hilton High School app.

This module intentionally has no application logic.  It only supplies the
card surface used by the main Python file so the data-processing behaviour
stays in one place and the UI can be adjusted independently.
"""

import tkinter as tk


class ModernCard(tk.Frame):
    """A lightweight card with a softly rounded-looking surface.

    Tkinter does not provide a native rounded Frame widget.  The corner masks
    below reveal the parent background at each corner while the normal Frame
    continues to accept pack/grid/place children exactly like the original
    LabelFrame did.
    """

    def __init__(self, parent, title="", bg="#ffffff", fg="#1e293b",
                 border="#e2e8f0", outside="#f1f5f9", **kwargs):
        self._outside = outside
        self._card_bg = bg
        self._border = border
        self._text = fg
        self._title_text = title
        super().__init__(
            parent, bg=bg, bd=0, relief="flat",
            highlightthickness=1, highlightbackground=border,
            highlightcolor=border, **kwargs
        )

        self._title_label = None
        if title:
            self._title_label = tk.Label(
                self, text=title, bg=bg, fg=fg,
                font=("Segoe UI", 9, "bold"), anchor="w"
            )
            self._title_label.pack(fill="x", pady=(0, 8))

        self._corner_masks = [
            tk.Frame(self, bg=outside, bd=0, highlightthickness=0),
            tk.Frame(self, bg=outside, bd=0, highlightthickness=0),
            tk.Frame(self, bg=outside, bd=0, highlightthickness=0),
            tk.Frame(self, bg=outside, bd=0, highlightthickness=0),
        ]
        self.bind("<Configure>", self._position_corner_masks, add="+")

    def _position_corner_masks(self, _event=None):
        size = 9
        width = max(self.winfo_width(), size)
        height = max(self.winfo_height(), size)
        positions = (
            (0, 0),
            (width - size, 0),
            (0, height - size),
            (width - size, height - size),
        )
        for mask, (x, y) in zip(self._corner_masks, positions):
            mask.place(x=x, y=y, width=size, height=size)
            mask.lift()

    def refresh_theme(self, outside, bg, border, fg):
        self._outside = outside
        self._card_bg = bg
        self._border = border
        self._text = fg
        self.configure(
            bg=bg, highlightbackground=border, highlightcolor=border)
        if self._title_label is not None:
            self._title_label.configure(bg=bg, fg=fg)
        for mask in self._corner_masks:
            mask.configure(bg=outside)