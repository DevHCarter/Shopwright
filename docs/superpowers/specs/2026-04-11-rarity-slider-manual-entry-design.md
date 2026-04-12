# Rarity Slider Manual Entry — Design Spec
**Date:** 2026-04-11

## Summary

Add manual text entry to each rarity percentage slider in the Shop Settings tab. The existing read-only label showing the current percentage value is replaced with an editable `tk.Entry` of the same size. Typing a value and pressing Enter (or tabbing away) applies it with the same auto-rebalancing logic used by the slider.

---

## Layout Change

Each rarity row currently has:
```
[Rarity Label (width=12)] [Slider (length=260, padx=8)] [% Label (width=5)]
```

After this change:
```
[Rarity Label (width=12)] [Slider (length=260, padx=8)] [Entry (width=4)]
```

The `tk.Label` (width=5, textvariable=lbl_var) is replaced with `tk.Entry` (width=4, textvariable=lbl_var). All other layout properties are unchanged.

### Entry styling
- `textvariable=lbl_var` — same StringVar already used by the old label
- `width=4`
- `font=("Consolas", 10)`
- `bg=c["bg"]`, `fg=<rarity color>`
- `relief="flat"`, `bd=1`
- `justify="right"`

---

## Interactions

### Slider → Entry (no change required)
`_on_slider` already sets `lbl_var` to `f"{var.get():>3}%"`. The entry reflects this automatically via the shared `StringVar`.

### Entry → Slider
A new method `_on_entry(rarity)` is bound to `<Return>` and `<FocusOut>` on each entry:

1. Read the current text from `lbl_var`, strip `%` and whitespace, attempt `int()` parse
2. On parse failure or empty string: reset `lbl_var` to the current slider value and return
3. Clamp the parsed int to `[0, 100]`
4. Call `self.rarity_sliders[rarity].set(clamped_value)`
5. Call `self._on_slider(rarity, str(clamped_value))` — reuses all existing rebalancing and label-refresh logic

### Invalid input handling
Any non-numeric or out-of-range input silently resets the entry to the current slider value. No error dialogs or validation messages.

---

## Files Changed

| File | Change |
|------|--------|
| `shop.py` | Replace `tk.Label` with `tk.Entry` in `_build_settings_tab` rarity loop (~line 2669). Add `_on_entry(rarity)` method near `_on_slider`. |

---

## Out of Scope
- No changes to slider length, row padding, total label, or reset button
- No changes to `_on_slider`, `_on_wealth_change`, or `_reset_distribution`
- No new state variables needed
