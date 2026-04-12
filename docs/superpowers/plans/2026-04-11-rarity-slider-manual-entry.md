# Rarity Slider Manual Entry Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Replace the read-only `% label` next to each rarity slider with an editable `tk.Entry` so users can type values directly, triggering the same auto-rebalancing logic used by the slider.

**Architecture:** Two surgical edits to `shop.py` — swap the `tk.Label` for a `tk.Entry` in `_build_settings_tab`, and add a `_on_entry` handler that parses/validates input and delegates to the existing `_on_slider` method.

**Tech Stack:** Python, Tkinter (`tk.Entry`, `tk.StringVar`)

---

### Task 1: Replace the read-only label with an editable Entry

**Files:**
- Modify: `shop.py:2669-2671`

- [ ] **Step 1: Replace the `tk.Label` widget with a `tk.Entry`**

In `_build_settings_tab`, inside the rarity loop, replace lines 2669–2671:

```python
            tk.Label(row, textvariable=lbl_var, width=5,
                     bg=c["bg"], fg=color,
                     font=("Consolas", 10)).pack(side="left")
```

with:

```python
            entry = tk.Entry(row, textvariable=lbl_var, width=4,
                             bg=c["bg"], fg=color,
                             font=("Consolas", 10),
                             relief="flat", bd=1,
                             justify="right",
                             insertbackground=color)
            entry.pack(side="left")
            entry.bind("<Return>",   lambda e, r=rarity: self._on_entry(r))
            entry.bind("<FocusOut>", lambda e, r=rarity: self._on_entry(r))
```

- [ ] **Step 2: Commit**

```bash
git add shop.py
git commit -m "feat: replace rarity % label with editable Entry widget"
```

---

### Task 2: Add the `_on_entry` handler

**Files:**
- Modify: `shop.py` — add method after `_on_slider` (~line 4014)

- [ ] **Step 1: Add `_on_entry` immediately after `_on_slider`**

Insert the following method after the closing line of `_on_slider` (after line 4013, before `def _on_wealth_change`):

```python
    def _on_entry(self, rarity: str):
        """Handle manual value typed into a rarity entry box.
        Parses the entry, clamps to [0, 100], then delegates to _on_slider
        so the same rebalancing logic applies."""
        raw = self.slider_labels[rarity].get().replace("%", "").strip()
        try:
            val = max(0, min(100, int(raw)))
        except ValueError:
            # Reset to current slider value on bad input
            self.slider_labels[rarity].set(
                f"{self.rarity_sliders[rarity].get():>3}%")
            return
        self.rarity_sliders[rarity].set(val)
        self._on_slider(rarity, str(val))
```

- [ ] **Step 2: Commit**

```bash
git add shop.py
git commit -m "feat: add _on_entry handler for rarity manual input"
```

---

### Task 3: Manual smoke test

No automated test infrastructure exists in this project. Verify the feature manually by running the app:

- [ ] **Step 1: Launch the app**

```bash
python shop.py
```

- [ ] **Step 2: Open the Settings tab and verify these cases**

| Action | Expected result |
|--------|----------------|
| Drag a rarity slider | Entry box updates to match new value |
| Click entry, type `50`, press Enter | Slider moves to 50; other sliders rebalance; Total stays 100% |
| Click entry, type `50`, click away (FocusOut) | Same as above |
| Type `999` and press Enter | Clamped to 100; other sliders go to 0; Total stays 100% |
| Type `abc` and press Enter | Entry resets to current slider value; no crash |
| Type `0` for all rarities except one | That one shows 100%; total stays 100% |
| Click "↺ Reset Distribution" | All entries reset to wealth preset values |
| Switch Wealth Level | All entries update to match preset |
