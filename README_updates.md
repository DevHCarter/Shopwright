# DND Shop Generator — Recent Updates

## Overview

Four major features were added to `DND_ShopGen.py`:

1. **Transaction / Purchase Log** — track items sold to players per session
2. **Light Mode** — a bright theme toggle in App Settings (⚙ icon, top right)
3. **General UI Improvements** — keyboard shortcuts, button tooltips, and a right-click context menu on the shop inventory
4. **App Settings consolidation** — Shop Mode, Table Display, Appearance, and Tab Visibility all moved to the ⚙ App Settings cog

---

## 1. Transaction / Purchase Log

### New Tab: 📜 Transaction Log

A new tab has been added to the far right of the tab bar. It shows a full sortable log of every item sold from any shop.

**Columns:** Timestamp | Shop | Item | Rarity | Qty | Price | Session | Action

**Filtering:**
- Filter by shop name using the "Shop:" dropdown
- Filter by session tag using the "Session:" dropdown

**Sorting:** Click any column header to sort ascending; click again for descending.

**Buttons:**
- **↻ Refresh** — manually reload the log from the database
- **Export CSV** — save the full log to a `.csv` file for external use
- **✖ Clear Log** — permanently delete all transaction records (confirms first)

The log tab auto-refreshes whenever you switch to it.

### Marking Items as Sold

Right-click any item in the shop inventory and choose **💰 Mark as Sold...**

A dialog will appear asking:
- **Quantity sold** (spinbox capped at the item's current stock)
- **Session tag** (optional free text, e.g. "Session 12" or "Session 5 - Waterdeep")

On confirm:
- The item's quantity is decremented in the shop inventory
- If quantity reaches 0, the item is removed from inventory
- A row is written to the `transactions` table in `shops.db`
- The status bar updates with a confirmation message

### Database

A new table `transactions` was added to `shop_data/shops.db`:

```sql
CREATE TABLE IF NOT EXISTS transactions (
    id          INTEGER PRIMARY KEY AUTOINCREMENT,
    shop_id     INTEGER,
    shop_name   TEXT,
    item_name   TEXT NOT NULL,
    rarity      TEXT,
    quantity    INTEGER DEFAULT 1,
    price       TEXT,
    session_tag TEXT DEFAULT '',
    action      TEXT DEFAULT 'sold',
    timestamp   TEXT DEFAULT (datetime('now'))
);
```

The table is created automatically on first launch via the existing migration pattern. No manual setup required.

---

## 2. Light Mode

### Toggling Light Mode

Click the **⚙** icon (top right) → **Appearance** section → check **Light Mode** → click **✔ Save**.

The entire UI instantly switches to a warm parchment-toned light theme:

| Element | Dark Mode | Light Mode |
|---------|-----------|------------|
| Background | `#1a1a2e` (dark navy) | `#f5f0e8` (parchment) |
| Text | `#e0d8c0` (cream) | `#2a2010` (dark brown) |
| Accent | `#c9a84c` (gold) | `#7a4f0e` (dark gold/brown) |
| Fields | `#2d2d4e` (dark blue) | `#e0d8c0` (light cream) |
| Header bars | `#0f0f1e` (near black) | `#ede5d0` (light parchment) |

Uncheck to return to dark mode at any time. All tabs, treeviews, and widgets update live — no restart required.

---

## 3. General UI Improvements

### Keyboard Shortcuts

| Shortcut | Action |
|----------|--------|
| `Ctrl+G` | Generate Shop |
| `Ctrl+R` | Reroll (10–30% of unlocked items) |
| `Ctrl+S` | Save Shop |
| `Ctrl+E` | Export Shop to JSON |
| `Space`  | Toggle lock on selected item (only fires when not typing in a field) |

### Button Tooltips

Hovering over the main action buttons in the Shop tab shows a tooltip with a brief description and the keyboard shortcut:

- ⚡ Generate Shop → `"Generate a new shop inventory  (Ctrl+G)"`
- ↻ Reroll → `"Replace ~20% of unlocked items  (Ctrl+R)"`
- ✖ Clear Shop → `"Clear the current shop inventory"`
- ＋ Add Item → `"Manually add an item from the database"`

### Right-click Context Menu

Right-clicking any row in the shop inventory opens a context menu with:

- **⟳ Reroll This Item** — replace the selected item with a fresh pick of the same rarity
- **◆ Toggle Lock** — lock or unlock the item (same as double-click)
- **💰 Mark as Sold...** — opens the sell dialog (see Transaction Log section above)

---

## 4. App Settings Cog (⚙ top right)

The App Settings window (opened via the ⚙ icon in the top-right bar) now contains all display and behaviour settings. The following sections were **moved from ⚙️ Stock Settings** into App Settings:

### Shop Mode
- **Non-Magical Shop** — restricts inventory to common items at most
- **DND Official Only** — removes all homebrew items (The Griffon's Saddlebag Books 1–5)

### Table Display
- **Show Quantity column** — toggle the Qty column in the shop inventory
- **Show DMG Price Range column** — toggle the estimated value column

### Appearance
- **Light Mode** — toggle between dark and light themes (see section 2 above)

### Tab Visibility
Control which optional tabs are shown in the main tab bar. The following three tabs are **always visible** and cannot be hidden:
- ▶ Shop
- ⚙️ Stock Settings
- ◆ Campaigns & Saves

The following tabs can be toggled on or off:

| Tab | Default |
|-----|---------|
| 💰 Sell Item | Shown |
| 🕮 Item Gallery | Shown |
| ✦ Shopkeeper | Shown |
| ℹ Shop Info | Shown |
| 📜 Transaction Log | Shown |

Changes take effect immediately when you click **✔ Save** in the App Settings dialog.

---

## Technical Notes

- All color logic is centralized in `_apply_theme(mode)` — adding new palette entries there is sufficient to extend both themes.
- The `ToolTip` class is a standalone utility defined just before `class ShopApp` and can be attached to any widget via `ToolTip(widget, "text")`.
- The `_recolor_widgets()` method walks the full widget tree recursively and updates all non-ttk widgets. TTK widget colors are handled by `ttk.Style` and update automatically.
- Row colors in all three treeviews (shop inventory, sell results, item gallery) now read from `self.ROW_ODD` / `self.ROW_EVEN` etc. instead of the module-level constants, so they respond to theme switches.
- Tab visibility uses `ttk.Notebook.hide()` / `ttk.Notebook.add()` to show/hide optional tabs while preserving their position and configuration.
