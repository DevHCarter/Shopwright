import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import csv
import json
import math
import sqlite3
import random
import re
from datetime import datetime
from pathlib import Path

# ── Quantity generation ────────────────────────────────────────────────────────

_TGS_SOURCES = {"TGS1", "TGS2", "TGS3", "TGS4", "TGS5"}

_VEHICLE_NAME_FRAGMENTS = {
    "ship", "galley", "longship", "keelboat", "rowboat", "warship",
    "whaleboat", "carriage", "wagon", "cart", "sled", "dogsled", "chariot",
}

def _is_vehicle(name: str) -> bool:
    return any(frag in name.lower() for frag in _VEHICLE_NAME_FRAGMENTS)

def _is_generic_variant(item: dict) -> bool:
    return ("Generic Variant" in item.get("Tags", "")
            or "Generic Variant" in item.get("Type", ""))

_SIZE_MOD_RANGES: dict[str, dict[str, tuple[float, float]]] = {
    "Village":    {"mundane": (0, 5),  "common": (0, 2),  "uncommon": (0, 0),
                   "rare":    (0, 0),  "very rare": (0, 0), "legendary": (0, 0)},
    "Town":       {"mundane": (0, 10), "common": (0, 4),  "uncommon": (0, 2),
                   "rare":    (0, 0),  "very rare": (0, 0), "legendary": (0, 0)},
    "City":       {"mundane": (2, 15), "common": (1, 5),  "uncommon": (1, 4),
                   "rare":    (0, 3),  "very rare": (0, 1), "legendary": (0, 0)},
    "Metropolis": {"mundane": (5, 30), "common": (3, 15), "uncommon": (2, 6),
                   "rare":    (0, 5),  "very rare": (0, 1), "legendary": (0, 0)},
}

def _get_size_mod(city_size: str, rarity: str) -> float:
    rarity_key = rarity.lower().strip()
    if rarity_key in ("none", ""):
        rarity_key = "mundane"
    if rarity_key in ("artifact", "varies", "unknown"):
        return 0.0
    size_table = _SIZE_MOD_RANGES.get(city_size, _SIZE_MOD_RANGES["Town"])
    lo, hi = size_table.get(rarity_key, (0, 0))
    return random.uniform(lo, hi) if hi > 0 else 0.0

def _get_item_weight(item: dict, tags: set[str]) -> int:
    """Return stackability weight (0 = singular/always qty 1, up to 3 = stacks heavily).

    Checks the Quantity column first as a manual override, then infers from
    rarity, source, tags, and item properties.
    """
    _CONSUMABLE_TAGS = {"Potion", "Scroll", "Ammunition", "Oil", "Dust/Powder", "Food/Drink"}
    rarity = item.get("Rarity", "").strip().lower()
    source = item.get("Source", "").strip()
    name   = item.get("Name",   "").lower()
    text   = item.get("Text",   "").lower()

    col_val = item.get("Quantity", "")
    if col_val and str(col_val).strip().isdigit():
        return int(str(col_val).strip())

    if rarity in ("legendary", "artifact"):          return 0
    if rarity == "very rare":                         return 0
    if source in _TGS_SOURCES and rarity == "rare":  return 0
    if "sentient" in text:                            return 0
    if _is_generic_variant(item):                     return 0
    if _is_vehicle(name):                             return 0

    if tags & _CONSUMABLE_TAGS:
        if rarity in ("mundane", "none", "common"):  return 3
        if rarity == "uncommon":                     return 2
        return 1 

    if rarity in ("mundane", "none"):  return 2
    if rarity == "common":             return 1
    if rarity == "uncommon":           return 1
    return 0

def generate_item_quantity(item: dict, city_size: str = "Town", wealth: str = "Average") -> int:
    """Qty = ceil((size_mod * weight) + 1), floored at 1.

    size_mod: random float from the city+rarity range table (0.0 when the
              rarity doesn't appear at that city size).
    weight:   stackability score 0-3 inferred from item properties.
    A weight of 0 always produces exactly 1 regardless of city size.
    """
    tags     = {t.strip() for t in item.get("Tags", "").split(",") if t.strip()}
    rarity   = item.get("Rarity", "").strip().lower()
    weight   = _get_item_weight(item, tags)
    size_mod = _get_size_mod(city_size, rarity)
    return max(1, math.ceil((size_mod * weight) + 1))

# ── Cultural tag filter ────────────────────────────────────────────────────────

CULTURAL_TAGS: set[str] = {
    "Draconic", "Drow", "Dwarven", "Elven", "Fey",
    "Fiendish", "Giant",
}

def culture_match(item: dict, active_culture: str | None) -> bool:
    """Return True if the item is compatible with the active culture filter.

    Rules:
      - No active culture  → everything passes.
      - Item has no cultural tag → Universal, always passes.
      - Item has a cultural tag  → passes only if it matches active_culture.
    """
    if not active_culture:
        return True
    item_tags = {t.strip() for t in item.get("Tags", "").split(",") if t.strip()}
    item_cultures = item_tags & CULTURAL_TAGS
    if not item_cultures:          # no cultural tag = Universal
        return True
    return active_culture in item_cultures


# ── Paths ──────────────────────────────────────────────────────────────────────
BASE_DIR  = Path(__file__).parent
DATA_DIR  = BASE_DIR / "shop_data"
DATA_DIR.mkdir(exist_ok=True)
DB_PATH   = DATA_DIR / "shops.db"

MASTER_CSV = BASE_DIR / "Items_Beta_1.csv"

SHOP_TYPE_TO_POOL = {
    "Alchemy":               "alchemy",
    "Armory":                "armory",
    "Blacksmith":            "blacksmith",
    "Fletcher & Bowyer":     "fletcher_bowyer",
    "General Store":         "general_store",
    "Jeweler & Curiosities": "jeweler",
    "Magic":                 "magic",
    "Scribe & Scroll":       "scribe_scroll",
    "Stables & Outfitter":   "stables",
    "Tavern & Inn":          "tavern",
}

SOURCE_BOOKS: dict[str, str] = {
    "AAG":      "AAG — Astral Adventurer's Guide",
    "AI":       "AI — Acquisitions Incorporated",
    "BAM":      "BAM — Boo's Astral Menagerie",
    "BGDIA":    "BGDIA — Baldur's Gate: Descent into Avernus",
    "BGG":      "BGG — Bigby Presents: Glory of the Giants",
    "BMT":      "BMT — Book of Many Things",
    "CM":       "CM — Candlekeep Mysteries",
    "CRCotN":   "CRCotN — Call of the Netherdeep",
    "CoA":      "CoA — Chains of Asmodeus",
    "CoS":      "CoS — Curse of Strahd",
    "DC":       "DC — Divine Contention",
    "DMG'14":   "DMG'14 — Dungeon Master's Guide 2014",
    "DMG'24":   "DMG'24 — Dungeon Master's Guide 2024",
    "DSotDQ":   "DSotDQ — Dragonlance: Shadow of the Dragon Queen",
    "DitLCoT":  "DitLCoT — Descent into the Lost Caverns",
    "EET":      "EET — Elemental Evil Supplement",
    "EFA":      "EFA — Eberron: From Aundair",
    "EGW":      "EGW — Explorer's Guide to Wildemount",
    "ERLW":     "ERLW — Eberron: Rising from the Last War",
    "FRAiF":    "FRAiF — FR: Against the Frostmaiden",
    "FRHoF":    "FRHoF — FR: Heroes of the Forgotten Realms",
    "FTD":      "FTD — Fizban's Treasury of Dragons",
    "GGR":      "GGR — Guildmasters' Guide to Ravnica",
    "GoS":      "GoS — Ghosts of Saltmarsh",
    "HftT":     "HftT — Hunt for the Thessalhydra",
    "HotB":     "HotB — Hoard of the Beast",
    "HotDQ":    "HotDQ — Hoard of the Dragon Queen",
    "IDRotF":   "IDRotF — Icewind Dale: Rime of the Frostmaiden",
    "JttRC":    "JttRC — Journeys through the Radiant Citadel",
    "KftGV":    "KftGV — Keys from the Golden Vault",
    "LFL":      "LFL — Lightning Fast",
    "LLK":      "LLK — Lost Laboratory of Kwalish",
    "LoX":      "LoX — Light of Xaryxis",
    "MM'14":    "MM'14 — Monster Manual 2014",
    "MOT":      "MOT — Mythic Odysseys of Theros",
    "MTF":      "MTF — Mordenkainen's Tome of Foes",
    "NF":       "NF — Netherdeep",
    "OotA":     "OotA — Out of the Abyss",
    "PHB'14":   "PHB'14 — Player's Handbook 2014",
    "PHB'24":   "PHB'24 — Player's Handbook 2024",
    "PaBTSO":   "PaBTSO — Phandelver and Below",
    "PotA":     "PotA — Princes of the Apocalypse",
    "QftIS":    "QftIS — Quests from the Infinite Staircase",
    "RMBRE":    "RMBRE — Rime of the Rimewind",
    "RoT":      "RoT — Rise of Tiamat",
    "RoTOS":    "RoTOS — Rise of Tiamat Online Supplement",
    "SCAG":     "SCAG — Sword Coast Adventurer's Guide",
    "SCC":      "SCC — Strixhaven: A Curriculum of Chaos",
    "SDW":      "SDW — Sleeping Dragon's Wake",
    "SKT":      "SKT — Storm King's Thunder",
    "SatO":     "SatO — Spelljammer: Adventures in Space",
    "TCE":      "TCE — Tasha's Cauldron of Everything",
    "TGS1":     "TGS1 — The Griffon's Saddlebag I",
    "TGS2":     "TGS2 — The Griffon's Saddlebag II",
    "TGS3":     "TGS3 — The Griffon's Saddlebag III",
    "TGS4":     "TGS4 — The Griffon's Saddlebag IV",
    "TGS5":     "TGS5 — The Griffon's Saddlebag V",
    "TftYP":    "TftYP — Tales from the Yawning Portal",
    "ToA":      "ToA — Tomb of Annihilation",
    "VEoR":     "VEoR — Vecna: Eve of Ruin",
    "VGM":      "VGM — Volo's Guide to Monsters",
    "VRGR":     "VRGR — Van Richten's Guide to Ravenloft",
    "WBtW":     "WBtW — The Wild Beyond the Witchlight",
    "WDH":      "WDH — Waterdeep: Dragon Heist",
    "WDMM":     "WDMM — Waterdeep: Dungeon of the Mad Mage",
    "WttHC":    "WttHC — Welcome to the Cynosure",
    "XGE":      "XGE — Xanathar's Guide to Everything",
}
_SOURCE_OPTS = ["(All)"] + [SOURCE_BOOKS[k] for k in sorted(SOURCE_BOOKS)]
RARITY_ORDER = {
    "mundane": 0, "none": 0, "common": 1, "uncommon": 2, "rare": 3,
    "very rare": 4, "legendary": 5, "artifact": 6,
    "varies": 7, "unknown": 8, "unknown (magic)": 9,
}

# ── Rarity display colours ─────────────────────────────────────────────────────
RARITY_COLORS_MAP: dict[str, str] = {
    "mundane":    "#999999",
    "none":       "#c8c8c8",
    "common":     "#c8c8c8",
    "uncommon":   "#1eff00",
    "rare":       "#0070dd",
    "very rare":  "#a335ee",
    "legendary":  "#ff8000",
    "artifact":   "#cc1212",
}

# ── DM's Guide price ranges ────────────────────────────────────────────────────
RARITY_PRICE_RANGES = {
    "mundane":         (1,      50),
    "none":            (1,      50),
    "common":          (50,     100),
    "uncommon":        (101,    500),
    "rare":            (501,    5000),
    "very rare":       (5001,   50000),
    "legendary":       (50001,  500000),
    "artifact":        (100000, 1000000),
    "varies":          (10,     500),
    "unknown":         (10,     100),
    "unknown (magic)": (50,     500),
}

MARKET_PRICE_RANGES: dict[str, tuple[int, int]] = {
    "common":    (50,    100),
    "uncommon":  (101,   500),
    "rare":      (501,   5000),
    "very rare": (5001,  50000),
    "legendary": (50001, 500000),
}

def generate_market_price(rarity: str) -> int | None:
    """Return a random integer market price for the given rarity, or None."""
    r = normalize_rarity(rarity)
    rng = MARKET_PRICE_RANGES.get(r)
    if rng is None:
        return None
    return random.randint(rng[0], rng[1])

# ── City size → item count ranges ─────────────────────────────────────────────
CITY_SIZE_RANGES = {
    "Village":    (10, 15),
    "Town":       (15, 25),
    "City":       (25, 35),
    "Metropolis": (35, 60),
}

# ── Wealth → rarity distribution ──────────────────────────────────────────────
WEALTH_DEFAULTS = {
    "Poor":    {"common": 55, "uncommon": 30, "rare": 15, "very rare": 0,  "legendary": 0,  "artifact": 0},
    "Average": {"common": 40, "uncommon": 30, "rare": 24, "very rare": 5,  "legendary": 1,  "artifact": 0},
    "Rich":    {"common": 30, "uncommon": 25, "rare": 20, "very rare": 15, "legendary": 10, "artifact": 0},
}

# ── Generative shop name parts ────────────────────────────────────────────────
# Six patterns are assembled at random from these pools:
#   A) "The [Adj] [Noun]"       e.g. "The Bubbling Cauldron"
#   B) "[Name]'s [Noun]"        e.g. "Aldric's Elixirs"
#   C) "The [Noun] & [Noun2]"   e.g. "The Quill & Candle"
#   D) "[Adj] [Trade]"          e.g. "Ironblood Smithy"
#   E) "[Name]'s [Trade]"       e.g. "Gornak's Forge"
#   F) "[Noun] & [Noun2]"       e.g. "Shield & Sword"
SHOP_NAME_PARTS: dict[str, dict[str, list[str]]] = {
    "Alchemy": {
        "adjectives":   ["Bubbling", "Smoky", "Gilded", "Cobalt", "Silver", "Amber",
                         "Dripping", "Misty", "Fizzling", "Sputtering", "Boiling",
                         "Leaking", "Crimson", "Verdant", "Acrid", "Fuming"],
        "nouns":        ["Cauldron", "Crucible", "Retort", "Vial", "Flask", "Mortar",
                         "Phial", "Alembic", "Tincture", "Concoction", "Burner", "Still"],
        "second_nouns": ["Bottle", "Smoke", "Powder", "Fume", "Extract", "Ember",
                         "Flame", "Vapour", "Ash"],
        "trade_words":  ["Apothecary", "Formulae", "Mixtures", "Remedies",
                         "Concoctions", "Philtres", "Elixirs", "Potions", "Distillery"],
        "npc_names":    ["Mira", "Aldric", "Yzara", "Seraphel", "Fizzwick",
                         "Madame Voss", "Thornwick", "Brimstone", "Ember", "Cobalt",
                         "Sable", "Orwick", "Fenrath"],
    },
    "Armory": {
        "adjectives":   ["Iron", "Brazen", "Tempered", "Steel", "Unyielding", "Dented",
                         "Cold", "Sunken", "Wyrm-Scale", "Gilded", "Forged", "Battle-Worn"],
        "nouns":        ["Bastion", "Bulwark", "Pauldron", "Visor", "Curtain",
                         "Guard", "Shell", "Plate", "Greave", "Vambrace"],
        "second_nouns": ["Sword", "Shield", "Mail", "Rivet", "Helm", "Buckler", "Hauberk"],
        "trade_words":  ["Armory", "Armaments", "Harness", "Outfitters", "Arms", "Works"],
        "npc_names":    ["Velthurin", "Harkon", "Ironforge", "Dragonsteel", "Aegis",
                         "Crestfall", "Coldmere", "Rampart", "Valdris", "Morthane"],
    },
    "Blacksmith": {
        "adjectives":   ["Red", "White-Hot", "Sooty", "Bent", "Clanging", "Ashen",
                         "Deepfire", "Glowing", "Cracked", "Hammered", "Scorched"],
        "nouns":        ["Anvil", "Hammer", "Forge", "Hearth", "Trough",
                         "Nail", "Spark", "Slag", "Bellows", "Tong"],
        "second_nouns": ["Flame", "Steel", "Cinder", "Ember", "Ash", "Coal", "Iron", "Blade"],
        "trade_words":  ["Smithy", "Forge", "Ironworks", "Smithworks",
                         "Metalworks", "Foundry", "Works"],
        "npc_names":    ["Gornak", "Halverson", "Embric", "Bram", "Stonemaul",
                         "Ironblood", "Deepfire", "Ashfall", "Thunderstrike",
                         "Durnok", "Heldra", "Korrund"],
    },
    "Fletcher & Bowyer": {
        "adjectives":   ["Singing", "Fletched", "Taut", "Notched", "Straight",
                         "Loosed", "Drawn", "Swift", "Silent", "Keen"],
        "nouns":        ["Quiver", "Arrow", "Stave", "String", "Nock",
                         "Bolt", "Shaft", "Wing", "Bow", "Fletch"],
        "second_nouns": ["Reed", "Feather", "Yew", "Goose-Feather", "Birch",
                         "Sinew", "Ash", "Maple"],
        "trade_words":  ["Archery", "Bowyers", "Fletchers", "Bowworks",
                         "Quarrels", "Arrowcraft"],
        "npc_names":    ["Elara", "Mirethil", "Farryn", "Silvan", "Windwhisper",
                         "Thornfield", "Ashwood", "Bramblewood", "Pinecroft",
                         "Sylvara", "Hethrin"],
    },
    "General Store": {
        "adjectives":   ["Dusty", "Packed", "Cluttered", "Overstuffed", "Reliable",
                         "Common", "Wandering", "Worn", "Trusty", "Humble"],
        "nouns":        ["Counter", "Satchel", "Shelf", "Stall", "Purse",
                         "Post", "Barrel", "Crate", "Rack"],
        "second_nouns": ["Rope", "Pack", "Goods", "Finds", "Wares", "Odds", "Ends"],
        "trade_words":  ["Mercantile", "Provisions", "Supplies", "Emporium",
                         "Wares", "Depot", "Trading Post", "General"],
        "npc_names":    ["Halvard", "Millbrook", "Thorngate", "Greymarsh",
                         "Briarvale", "Cobblestone", "Dunmore", "Aldwick",
                         "Ferris", "Hadley"],
    },
    "Jeweler & Curiosities": {
        "adjectives":   ["Gilded", "Hidden", "Polished", "Glinting", "Shining",
                         "Whispering", "Lustrous", "Peculiar", "Gleaming", "Veiled"],
        "nouns":        ["Gem", "Jewel", "Stone", "Cabinet", "Hoard",
                         "Cache", "Trinket", "Find", "Vault", "Eye"],
        "second_nouns": ["Onyx", "Opal", "Varnish", "Luster", "Pearl",
                         "Sapphire", "Garnet", "Crystal", "Amber"],
        "trade_words":  ["Jewellers", "Gems", "Treasures", "Curios",
                         "Oddments", "Ornaments", "Antiquities"],
        "npc_names":    ["Tindra", "Aurelius", "Crystalveil", "Silverthread",
                         "Velvet", "Opalvane", "Lumen", "Gemwright",
                         "Sorra", "Nilvaris"],
    },
    "Magic": {
        "adjectives":   ["Arcane", "Enchanted", "Glowing", "Runic", "Umbral",
                         "Veilborn", "Starlit", "Bound", "Drifting", "Mystic",
                         "Wandering", "Whispering", "Sigil-Touched"],
        "nouns":        ["Attic", "Cache", "Shelf", "Grimoire", "Tome",
                         "Pocket", "Circle", "Vestibule", "Nook", "Alcove"],
        "second_nouns": ["Seal", "Sorcery", "Mist", "Veil", "Rune",
                         "Sigil", "Cantrip", "Aether", "Void"],
        "trade_words":  ["Magicka", "Arcana", "Enchantments", "Wares",
                         "Curios", "Magics", "Emporium", "Curiosities"],
        "npc_names":    ["Mystara", "Vexor", "Elara", "Mirrorgate", "Aethermist",
                         "Umbral", "Stardust", "Veilborn", "Zephyra",
                         "Mordecai", "Thessaly", "Ilvar"],
    },
    "Scribe & Scroll": {
        "adjectives":   ["Blotted", "Dusty", "Sealed", "Pressed", "Open",
                         "Careful", "Illumined", "Faded", "Inked", "Worn"],
        "nouns":        ["Quill", "Feather", "Folio", "Seal", "Hand",
                         "Letter", "Page", "Tome", "Scroll", "Script"],
        "second_nouns": ["Parchment", "Candle", "Vellum", "Ink", "Sigil",
                         "Wax", "Reed", "Ribbon", "Clasp"],
        "trade_words":  ["Scrivenery", "Transcripts", "Scrollworks",
                         "Scripts", "Manuscripts", "Calligraphy", "Bindery"],
        "npc_names":    ["Aldenmoor", "Thornwick", "Pencraft", "Reedham",
                         "Memoranda", "Vellum", "Inksworth", "Quillsby",
                         "Harrold", "Cressida"],
    },
    "Stables & Outfitter": {
        "adjectives":   ["Dusty", "Muddy", "Padded", "Tired", "Open",
                         "Stamping", "Cobbled", "Worn", "Weathered"],
        "nouns":        ["Hoof", "Saddle", "Shoe", "Paddock", "Yard",
                         "Spur", "Gate", "Stable", "Post"],
        "second_nouns": ["Tack", "Trail", "Feed", "Bridle", "Stirrup",
                         "Harness", "Rein", "Mane"],
        "trade_words":  ["Stables", "Livery", "Outfitters",
                         "Equestrian", "Mounts", "Feed & Tack"],
        "npc_names":    ["Ironmane", "Farrow", "Thunderhoof", "Crossroads",
                         "Briarvale", "Greystream", "Cloverfield", "Stonepath",
                         "Willowmere", "Crestfall", "Dusthoof", "Mirren"],
    },
    "Tavern & Inn": {
        "adjectives":   ["Rusty", "Golden", "Broken", "Tipsy", "Stumbling",
                         "Forgotten", "Half-Empty", "Laughing", "Sleeping",
                         "Salted", "Warm", "Last", "Crooked", "Drunken"],
        "nouns":        ["Flagon", "Goose", "Compass", "Cup", "Bard",
                         "Hearth", "Lantern", "Skull", "Dragon", "Boot",
                         "Griffon", "Boar", "Hound", "Raven"],
        "second_nouns": ["Moon", "Mead", "Wound", "Coin", "Candle",
                         "Ember", "Barrel", "Sword", "Pipe"],
        "trade_words":  ["Inn", "Tavern", "Rest", "Lodge",
                         "Alehouse", "Boarding House", "Roadhouse"],
        "npc_names":    ["Widow Harken", "Old Marten", "Crossroads",
                         "Hearthstone", "Dunwall", "Gretta", "Tobbin",
                         "Mirla", "Aldous", "Fenwick"],
    },
}

_NAME_PATTERNS = [
    "the_adj_noun",      
    "name_noun",         
    "the_noun_and_noun", 
    "adj_trade",        
    "name_trade",        
    "noun_and_noun",     
]

def generate_shop_name(shop_type: str) -> str:
    """Assemble a shop name from parts using a random structural pattern."""
    parts = SHOP_NAME_PARTS.get(shop_type)
    if not parts:
        return f"The {shop_type} Shop"

    pattern  = random.choice(_NAME_PATTERNS)
    adj      = random.choice(parts["adjectives"])
    noun     = random.choice(parts["nouns"])
    trade    = random.choice(parts["trade_words"])
    npc      = random.choice(parts["npc_names"])
    all_nouns = parts["nouns"] + parts["second_nouns"]
    n1       = random.choice(all_nouns)
    n2       = random.choice([n for n in all_nouns if n != n1] or all_nouns)

    if pattern == "the_adj_noun":
        return f"The {adj} {noun}"
    elif pattern == "name_noun":
        return f"{npc}'s {noun}"
    elif pattern == "the_noun_and_noun":
        return f"The {n1} & {n2}"
    elif pattern == "adj_trade":
        return f"{adj} {trade}"
    elif pattern == "name_trade":
        return f"{npc}'s {trade}"
    else:
        return f"{n1} & {n2}"


# ── Shopkeeper Generator ──────────────────────────────────────────────────────
SHOPKEEPER_POOLS: dict[str, list[str]] = {
    "races": [
        "Human", "Dwarf", "Elf", "Half-Elf", "Half-Orc", "Gnome",
        "Halfling", "Tiefling", "Dragonborn", "Goliath", "Aasimar",
        "Tabaxi", "Kenku", "Lizardfolk", "Tortle", "Firbolg",
        "Shadar-kai", "Sea Elf", "Wood Elf", "Deep Gnome",
    ],
    "personalities": [
        "Gruff but fair — few words, honest prices",
        "Overly cheerful — smiles through every haggle",
        "Paranoid and suspicious — eyes the door constantly",
        "Warm and motherly — treats every customer like kin",
        "Sarcastic and witty — a quip for every question",
        "Absent-minded — frequently loses track of their own sentences",
        "Deeply proud of their craft — lectures if given the chance",
        "Hates small talk — get to the point or get out",
        "Gossipy — knows every rumour in town",
        "Mournful — carries some old grief they won't explain",
        "Boastful — exaggerates every story dramatically",
        "Nervous — startles easily and speaks too fast",
        "Philosophical — turns any transaction into a meditation",
        "Blunt to the point of rudeness — but means well",
        "Theatrical — describes every item like a market-day barker",
        "Exhausted — running this shop alone and showing it",
        "Calculating — mentally tallying the profit of every word",
        "Curious — peppers customers with questions about their travels",
        "Devout — quietly prays before each transaction",
        "Former adventurer — has strong opinions on everything you pick up",
    ],
    "appearances": [
        "Broad-shouldered with calloused, ink-stained hands",
        "Slight and quick-moving, always straightening something",
        "Tall with sharp eyes that miss nothing",
        "Short and round, perpetually dusted with their trade",
        "Weathered face, grey-streaked hair pulled back tightly",
        "Impeccably dressed despite the grimy surroundings",
        "Covered in old burn scars they never mention",
        "Missing two fingers on the left hand",
        "Wears an excessive number of rings",
        "Smells strongly of their trade's particular materials",
        "Has a prominent scar across the bridge of the nose",
        "Remarkably young-looking for someone who's 'been here forty years'",
        "Squinting — needs spectacles but refuses to admit it",
        "Hair stands up in wild directions no matter what they do",
        "Immaculately groomed, not a hair out of place",
        "Tattooed arms visible beneath rolled-up sleeves",
        "Wears a coin on a cord around the neck — never explains it",
        "One eye is glass, the other studies you very carefully",
        "Moves with an old injury — favours the left leg",
        "Always has a mug of something nearby, perpetually half-drunk",
    ],
    "quirks": [
        "Hums tunelessly while they work",
        "Refuses to sell anything on a Tuesday",
        "Keeps a small shrine to a forgotten deity behind the counter",
        "Writes everything down in a tiny leather journal",
        "Absolutely hates rats — will stop mid-sentence if one appears",
        "Insists on sealing every deal with a handshake",
        "Calls every customer 'friend' regardless of history",
        "Never gives change — rounds prices in their favour, always",
        "Has a cat that sits on the counter and judges everyone",
        "Speaks in a heavy regional accent they've never lost",
        "Recites prices in rhyming couplets when in a good mood",
        "Refuses to discuss where certain items came from",
    ],
}

SHOPKEEPER_FIRST_NAMES: list[str] = [
    "Aldric", "Mira", "Gornak", "Elara", "Thornwick", "Sable",
    "Fenrath", "Cressida", "Velthurin", "Halvard", "Tindra",
    "Aurelius", "Mystara", "Vexor", "Gretta", "Tobbin", "Mirla",
    "Aldous", "Bram", "Heldra", "Durnok", "Sylvara", "Farryn",
    "Odalys", "Rennick", "Isolde", "Cassian", "Velka", "Storn",
    "Wren", "Lukas", "Petra", "Aldwyn", "Serith", "Naeris",
]
SHOPKEEPER_SURNAMES: list[str] = [
    "Ironforge", "Ashfall", "Briarvale", "Gemwright", "Coldmere",
    "Quillsby", "Copperkettle", "Dusthoof", "Hearthstone", "Varnish",
    "Dunwall", "Opalvane", "Thorngate", "Millbrook", "Stonemaul",
    "Inksworth", "Greymarsh", "Windwhisper", "Bramblewood", "Crestfall",
]

def generate_shopkeeper(shop_type: str = "") -> dict[str, str]:
    """Return a dict with keys: name, race, personality, appearance, quirk."""
    first = random.choice(SHOPKEEPER_FIRST_NAMES)
    last  = random.choice(SHOPKEEPER_SURNAMES)
    return {
        "name":        f"{first} {last}",
        "race":        random.choice(SHOPKEEPER_POOLS["races"]),
        "personality": random.choice(SHOPKEEPER_POOLS["personalities"]),
        "appearance":  random.choice(SHOPKEEPER_POOLS["appearances"]),
        "quirk":       random.choice(SHOPKEEPER_POOLS["quirks"]),
    }


# ── Tag filter categories ──────────────────────────────────────────────────────
TAG_CATEGORIES = {
    "Race/Creature": [
        "Drow", "Draconic", "Dwarven", "Elven", "Fey", "Fiendish", "Giant",
    ],
    "Damage/Element": [
        "Acid", "Fire", "Force", "Ice/Cold", "Lightning", "Necrotic",
        "Poison", "Psychic", "Radiant", "Thunder", "Slashing", "Piercing", "Bludgeoning",
    ],
    "Item Slot/Form": [
        "Adventuring Gear", "Ammunition", "Artisans", "Tools", "Amulet/Necklace",
        "Belt", "Book/Tome", "Boots/Footwear", "Card/Deck", "Cloak",
        "Dust/Powder", "Figurine", "Food/Drink", "Gloves/Bracers", "Headwear",
        "Instrument", "Potion", "Ring", "Rod", "Scroll", "Staff", "Tattoo",
        "Wand", "Other", "Trade Good", "Spellcasting Focus",
    ],
    "Weapon & Armor": [
        "Armor", "Finesse", "Generic Variant", "Heavy Armor", "Heavy Weapon",
        "Light Armor", "Light Weapon", "Medium Armor", "Melee", "Ranged Weapon",
        "Shield", "Thrown", "Two-Handed", "Versatile", "Weapon",
    ],
    "Rarity": [
        "Artifact", "Common", "Legendary", "Mundane", "Rare", "Uncommon", "Very Rare",
    ],
}

# ── Alternating row palette ────────────────────────────────────────────────────
ROW_ODD          = "#1e1e30"
ROW_EVEN         = "#171725"
ROW_LOCKED_ODD   = "#1e1e10"
ROW_LOCKED_EVEN  = "#17170d"
ROW_SELECTED     = "#2e2a14"


#  ─ Helpers ────────────────────────────────────────────────────
def normalize_rarity(r: str) -> str:
    return (r or "").strip().lower()

def rarity_rank(r: str) -> int:
    return RARITY_ORDER.get(normalize_rarity(r), 99)

def parse_given_cost(value_str: str) -> float | None:
    if not value_str:
        return None
    s = value_str.upper().replace(",", "")
    m = re.search(r"([\d.]+)\s*(GP|SP|CP)", s)
    if not m:
        return None
    amount = float(m.group(1))
    unit = m.group(2)
    if unit == "SP": amount /= 10
    elif unit == "CP": amount /= 100
    return round(amount, 2)

def weighted_rarity_pick(weights: dict[str, int]) -> str:
    pool: list[str] = []
    for rarity, pct in weights.items():
        pool.extend([rarity] * pct)
    while len(pool) < 100:
        pool.append("common")
    return random.choice(pool)

def format_currency(gp_value) -> str:
    """Convert a GP float to a multi-denomination display string.

    Internally everything is stored and calculated in GP (float).
    This converts to the smallest necessary denominations for display:
      1 gp = 10 sp = 100 cp

    Examples:
      15.0   → "15 gp"
      1.5    → "1 gp 5 sp"
      0.5    → "5 sp"
      0.07   → "7 cp"
      12.34  → "12 gp 3 sp 4 cp"
      0.0    → "—"
    """
    if gp_value is None or gp_value == "":
        return "—"
    try:
        total_cp = round(float(gp_value) * 100)
    except (TypeError, ValueError):
        return "—"
    if total_cp <= 0:
        return "—"
    gp = total_cp // 100
    sp = (total_cp % 100) // 10
    cp = total_cp % 10
    parts = []
    if gp: parts.append(f"{gp:,} gp")
    if sp: parts.append(f"{sp} sp")
    if cp: parts.append(f"{cp} cp")
    return " ".join(parts) if parts else "—"

def parse_pipe_table(table_str: str) -> list[dict]:
    """Parse a pipe-delimited table string from the CSV Table column.

    Returns a list of table dicts, each with:
        {"headers": [...], "rows": [[...], ...]}

    Multiple tables in one string are separated by a blank line.
    Returns an empty list if table_str is empty or invalid.
    """
    if not table_str or not table_str.strip():
        return []

    blocks = re.split(r'\n\s*\n', table_str.strip())
    result = []
    for block in blocks:
        lines = [ln.strip() for ln in block.strip().splitlines() if ln.strip()]
        if not lines:
            continue
        parsed = [[cell.strip() for cell in ln.split("|")] for ln in lines]
        headers = parsed[0]
        rows    = parsed[1:]
        if headers:
            result.append({"headers": headers, "rows": rows})
    return result



def apply_price_mod(cost_str: str, mod: int) -> str:
    if mod == 100 or not cost_str or cost_str == "—":
        return cost_str or "—"
    val = parse_given_cost(cost_str)
    if val is None:
        return cost_str
    modified = max(0.01, val * mod / 100)   # floor at 1 cp (0.01 gp)
    return format_currency(modified)


#  ── Database ────────────────────────────────────────────────────
def init_db():
    con = sqlite3.connect(DB_PATH)
    con.execute("PRAGMA foreign_keys = ON")
    cur = con.cursor()
    cur.executescript("""
        CREATE TABLE IF NOT EXISTS campaigns (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            name TEXT NOT NULL UNIQUE,
            created_at TEXT DEFAULT (datetime('now'))
        );
        CREATE TABLE IF NOT EXISTS towns (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            campaign_id INTEGER REFERENCES campaigns(id) ON DELETE CASCADE,
            name TEXT NOT NULL,
            city_size TEXT,
            created_at TEXT DEFAULT (datetime('now'))
        );
        CREATE TABLE IF NOT EXISTS shops (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            town_id INTEGER REFERENCES towns(id) ON DELETE CASCADE,
            name TEXT NOT NULL,
            shop_type TEXT,
            wealth TEXT,
            last_restocked TEXT,
            created_at TEXT DEFAULT (datetime('now'))
        );
        CREATE TABLE IF NOT EXISTS shop_items (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            shop_id INTEGER REFERENCES shops(id) ON DELETE CASCADE,
            item_id TEXT,
            name TEXT,
            rarity TEXT,
            item_type TEXT,
            source TEXT,
            page TEXT,
            cost_given TEXT,
            quantity TEXT,
            locked INTEGER DEFAULT 0,
            attunement TEXT,
            damage TEXT,
            properties TEXT,
            mastery TEXT,
            weight TEXT,
            tags TEXT,
            description TEXT,
            table_data TEXT DEFAULT ''
        );
    """)
    con.commit()

    try:
        con.execute("ALTER TABLE shops ADD COLUMN notes TEXT DEFAULT ''")
        con.commit()
    except sqlite3.OperationalError:
        pass
    try:
        con.execute("ALTER TABLE shop_items ADD COLUMN table_data TEXT DEFAULT ''")
        con.commit()
    except sqlite3.OperationalError:
        pass
    try:
        con.execute("ALTER TABLE shop_items ADD COLUMN sane_cost TEXT DEFAULT ''")
        con.commit()
    except sqlite3.OperationalError:
        pass
    try:
        con.execute("ALTER TABLE shop_items ADD COLUMN market_price TEXT DEFAULT ''")
        con.commit()
    except sqlite3.OperationalError:
        pass
    try:
        con.execute("ALTER TABLE shops ADD COLUMN shopkeeper_name TEXT DEFAULT ''")
        con.commit()
    except sqlite3.OperationalError:
        pass
    try:
        con.execute("ALTER TABLE shops ADD COLUMN shopkeeper_race TEXT DEFAULT ''")
        con.commit()
    except sqlite3.OperationalError:
        pass
    try:
        con.execute("ALTER TABLE shops ADD COLUMN shopkeeper_personality TEXT DEFAULT ''")
        con.commit()
    except sqlite3.OperationalError:
        pass
    try:
        con.execute("ALTER TABLE shops ADD COLUMN shopkeeper_appearance TEXT DEFAULT ''")
        con.commit()
    except sqlite3.OperationalError:
        pass
    con.close()


#  ── Item Loading ────────────────────────────────────────────────────

ALL_ITEMS: dict[str, list[dict]] = {}   # pool_key → [item, ...]
ALL_ITEMS_FLAT: list[dict] = []         # all items for sell lookup

def load_all_items():
    global ALL_ITEMS, ALL_ITEMS_FLAT
    if not MASTER_CSV.exists():
        print(f"[ERROR] Master CSV not found: {MASTER_CSV}")
        return

    pool_buckets: dict[str, list[dict]] = {
        pool_key: [] for pool_key in SHOP_TYPE_TO_POOL.values()
    }
    all_flat: list[dict] = []

    with open(MASTER_CSV, newline="", encoding="utf-8-sig") as f:
        for row in csv.DictReader(f):
            row = {k: (v.strip() if isinstance(v, str) else "") for k, v in row.items()}
            all_flat.append(row)
            pools = [p.strip() for p in row.get("Pools", "").split("|") if p.strip()]
            for pool_key in pools:
                if pool_key in pool_buckets:
                    pool_buckets[pool_key].append(row)

    pool_to_display = {v: k for k, v in SHOP_TYPE_TO_POOL.items()}
    for pool_key, items in pool_buckets.items():
        display_name = pool_to_display.get(pool_key, pool_key)
        ALL_ITEMS[display_name] = items

    ALL_ITEMS_FLAT.extend(all_flat)
    print(f"[INFO] Loaded {len(all_flat)} items from master CSV.")
    for display, items in sorted(ALL_ITEMS.items(), key=lambda x: -len(x[1])):
        print(f"         {display}: {len(items)} items")


#  ─── Shop Generation ────────────────────────────────────────────────────
def generate_shop_items(
    shop_type: str,
    count: int,
    rarity_weights: dict[str, int],
    existing_locked: list[dict] | None = None,
    tag_filters: set[str] | None = None,
    tag_excludes: set[str] | None = None,
    city_size: str = "Town",
    wealth: str = "Average",
    culture: str | None = None,
    mundane_only: bool = False,
) -> list[dict]:
    if shop_type not in ALL_ITEMS or not ALL_ITEMS[shop_type]:
        return []

    source_items = ALL_ITEMS[shop_type]
    buckets: dict[str, list[dict]] = {}
    for item in source_items:
        r = normalize_rarity(item.get("Rarity", "mundane"))
        buckets.setdefault(r, []).append(item)

    locked_items   = existing_locked or []
    locked_names   = {i["name"] for i in locked_items}
    needed         = count - len(locked_items)
    if needed <= 0:
        return locked_items

    generated      = list(locked_items)
    existing_names = set(locked_names)
    attempts       = 0
    fallback_order = ["mundane", "common", "uncommon", "rare", "none", "very rare", "legendary"]

    def tag_match(item: dict) -> bool:
        """Exclude beats include. Excluded tags hard-block; include filters
        then require at least one match (OR logic). No filters = allow all.
        Culture filter: items without any cultural tag are Universal (always
        pass); items with a cultural tag only pass if it matches active culture."""
        item_tags = {t.strip() for t in item.get("Tags", "").split(",") if t.strip()}
        if tag_excludes and (item_tags & tag_excludes):
            return False
        if tag_filters and not (item_tags & tag_filters):
            return False
        if not culture_match(item, culture):
            return False
        if mundane_only:
            r = normalize_rarity(item.get("Rarity", ""))
            if r not in ("mundane", "none", "common", ""):
                return False
        return True

    while len(generated) - len(locked_items) < needed and attempts < needed * 20:
        attempts += 1
        rarity = weighted_rarity_pick(rarity_weights)
        chosen_item = None
        for r in [rarity] + [x for x in fallback_order if x != rarity]:
            bucket    = buckets.get(r, [])
            available = [x for x in bucket
                         if x["Name"] not in existing_names and tag_match(x)]
            if available:
                chosen_item = random.choice(available)
                break
        if not chosen_item:
            continue

        cost_given = chosen_item.get("Value", "")
        quantity   = str(generate_item_quantity(chosen_item, city_size, wealth))
        rarity_raw = chosen_item.get("Rarity", "")
        mkt = generate_market_price(rarity_raw)

        generated.append({
            "item_id":      chosen_item.get("Item ID", ""),
            "name":         chosen_item.get("Name", ""),
            "rarity":       rarity_raw,
            "item_type":    chosen_item.get("Type", ""),
            "source":       chosen_item.get("Source", ""),
            "page":         chosen_item.get("Page", ""),
            "cost_given":   cost_given,
            "quantity":     quantity,
            "locked":       False,
            "attunement":   chosen_item.get("Attunement", ""),
            "damage":       chosen_item.get("Damage", ""),
            "properties":   chosen_item.get("Properties", ""),
            "mastery":      chosen_item.get("Mastery", ""),
            "weight":       chosen_item.get("Weight", ""),
            "tags":         chosen_item.get("Tags", ""),
            "description":  chosen_item.get("Text", ""),
            "table_data":   chosen_item.get("Table", ""),
            "sane_cost":    chosen_item.get("Sane_Cost", ""),
            "market_price": str(mkt) if mkt is not None else "",
        })
        existing_names.add(chosen_item["Name"])

    return generated


# ── GUI ────────────────────────────────────────────────────
class ShopApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("D&D Shop Generator")
        self.geometry("1380x860")
        self.minsize(1100, 700)
        self.configure(bg="#1a1a2e")
        self._apply_theme()

        # ── State ────────────────────────────────────────────────────
        self.current_items: list[dict] = []
        self.current_shop_type = tk.StringVar(value="Magic")
        self.city_size_var     = tk.StringVar(value="Town")
        self.wealth_var        = tk.StringVar(value="Average")
        self.culture_var       = tk.StringVar(value="")   # "" = no culture filter
        self.shop_name_var     = tk.StringVar(value="")
        self.selected_row      = None
        self._sort_col         = "rarity"
        self._sort_asc         = True
        self.price_modifier    = tk.IntVar(value=100)
        self._inspect_expanded = False   # whether inspector is in focus mode
        self.mundane_only_var  = tk.BooleanVar(value=False)  # mundane-only mode

        # Table display column toggles
        self.show_qty_col      = tk.BooleanVar(value=True)
        self.show_est_val_col  = tk.BooleanVar(value=True)

        # Rarity slider vars
        self.rarity_sliders: dict[str, tk.IntVar] = {
            r: tk.IntVar(value=v)
            for r, v in WEALTH_DEFAULTS["Average"].items()
        }

    
        self.active_tag_filters:   set[str] = set()
        self.excluded_tag_filters: set[str] = set()
        self._tag_state_vars:      dict[str, tk.IntVar] = {}


        self.sell_search_var    = tk.StringVar()
        self.sell_pct_var       = tk.IntVar(value=80)
        self.sell_selected_item = None 
        self._sell_popup        = None   
        self.shop_notes_widget  = None   

        self.shopkeeper_name_var        = tk.StringVar(value="")
        self.shopkeeper_race_var        = tk.StringVar(value="")
        self.shopkeeper_personality_var = tk.StringVar(value="")
        self.shopkeeper_appearance_var  = tk.StringVar(value="")

        self._custom_races: list[str] = list(SHOPKEEPER_POOLS["races"])

        self._app_settings = {
            "default_shop_type":   "Magic",
            "default_city_size":   "Town",
            "default_wealth":      "Average",
            "default_price_mod":   100,
            "auto_name_on_change": True,
            "gallery_per_page":    250,
        }

        self._gallery_sort_col     = "name"
        self._gallery_sort_asc     = True
        self._gallery_results:     list[dict] = []
        self._gallery_all_results: list[dict] = []
        self._gallery_page         = 0
        self._gallery_per_page     = tk.IntVar(value=250)

        init_db()
        load_all_items()
        self._build_ui()
        self._refresh_campaign_list()

    # ── Theme ─────────────────────────────────────────────────────────────────
    def _apply_theme(self):
        style = ttk.Style(self)
        style.theme_use("clam")
        bg, fg, accent, sel = "#1a1a2e", "#e0d8c0", "#c9a84c", "#2d2d4e"
        hdr = "#0f0f1e"

        style.configure(".",           background=bg, foreground=fg, font=("Georgia", 10))
        style.configure("TNotebook",   background=hdr, borderwidth=0)
        style.configure("TNotebook.Tab", background=sel, foreground=fg,
                        padding=[14, 6], font=("Georgia", 10, "bold"))
        style.map("TNotebook.Tab",
                  background=[("selected", accent)],
                  foreground=[("selected", hdr)])
        style.configure("TFrame",  background=bg)
        style.configure("TLabel",  background=bg, foreground=fg)
        style.configure("TButton", background=accent, foreground=hdr,
                        font=("Georgia", 10, "bold"), padding=6, relief="flat")
        style.map("TButton", background=[("active", "#e6c06a")])
        style.configure("Danger.TButton", background="#8b0000", foreground="#e0d8c0")
        style.map("Danger.TButton", background=[("active", "#b22222")])
        style.configure("Treeview",
                        background=ROW_ODD, foreground=fg,
                        fieldbackground=ROW_ODD, rowheight=26,
                        font=("Consolas", 9))
        style.configure("Treeview.Heading",
                        background=hdr, foreground=accent,
                        font=("Georgia", 9, "bold"))
        style.map("Treeview",
                  background=[("selected", accent)],
                  foreground=[("selected", hdr)])
        style.configure("TCombobox", fieldbackground=sel, background=sel, foreground=fg)
        style.configure("TScale",  background=bg, troughcolor=sel)
        style.configure("TEntry",  fieldbackground=sel, foreground=fg, insertcolor=fg)
        style.configure("TSeparator", background=accent)
        self.colors = {"bg": bg, "fg": fg, "accent": accent, "sel": sel, "hdr": hdr}

    # ── Main UI ───────────────────────────────────────────────────────────────
    def _build_ui(self):
        c = self.colors

        # ── Top bar
        top = tk.Frame(self, bg=c["hdr"], pady=6)
        top.pack(fill="x")

        tk.Label(top, text="[ D&D Shop Generator ]",
                 font=("Georgia", 15, "bold"),
                 bg=c["hdr"], fg=c["accent"]).pack(side="left", padx=14)

        tk.Label(top, text="Name:", bg=c["hdr"], fg=c["fg"],
                 font=("Georgia", 10)).pack(side="left", padx=(0, 4))
        tk.Entry(top, textvariable=self.shop_name_var, width=26,
                 bg=c["sel"], fg=c["fg"], insertbackground=c["fg"],
                 relief="flat", font=("Georgia", 10)).pack(side="left", padx=(0, 8))

        ttk.Button(top, text="↻ Name",
                   command=self._random_name).pack(side="left", padx=(0, 16))

        # ── Gear / App Settings button (top-right) 
        gear_btn = tk.Label(top, text="⚙", font=("Georgia", 16),
                            bg=c["hdr"], fg=c["fg"], cursor="hand2", padx=10)
        gear_btn.pack(side="right", padx=(0, 8))
        gear_btn.bind("<Button-1>", lambda e: self._open_app_settings_window())
        gear_btn.bind("<Enter>", lambda e: gear_btn.configure(fg=c["accent"]))
        gear_btn.bind("<Leave>", lambda e: gear_btn.configure(fg=c["fg"]))
        tk.Label(top, text="App Settings", font=("Georgia", 8, "italic"),
                 bg=c["hdr"], fg=c["fg"], cursor="hand2").pack(side="right")

        # ── Notebook ───────────────────────────────────────────────────────────
        nb = ttk.Notebook(self)
        nb.pack(fill="both", expand=True, padx=8, pady=(4, 8))

        self.tab_action      = ttk.Frame(nb)
        self.tab_settings    = ttk.Frame(nb)
        self.tab_sell        = ttk.Frame(nb)
        self.tab_save        = ttk.Frame(nb)
        self.tab_gallery     = ttk.Frame(nb)
        self.tab_shopkeeper  = ttk.Frame(nb)
        nb.add(self.tab_action,      text="  ⚡ Action  ")
        nb.add(self.tab_settings,    text="  ⚙ Stock Settings  ")
        nb.add(self.tab_sell,        text="  ◈ Sell Item  ")
        nb.add(self.tab_save,        text="  ◆ Campaigns & Saves  ")
        nb.add(self.tab_gallery,     text="  ◉ Item Gallery  ")
        nb.add(self.tab_shopkeeper,  text="  ✦ Shopkeeper  ")

        self._build_action_tab()
        self._build_settings_tab()
        self._build_sell_tab()
        self._build_save_tab()
        self._build_gallery_tab()
        self._build_shopkeeper_tab()
        self._setup_hover_scroll()

    # ── Global hover-aware scroll ─────────────────────────────────────────────
    def _setup_hover_scroll(self):
        """Single app-level MouseWheel handler — scrolls whichever scrollable
        widget the cursor is hovering over regardless of focus.

        Handles: Canvas, Text, Listbox, Treeview
        Platforms: Windows (delta ±120), macOS (delta ±3–5), Linux (Button-4/5)
        """
        SCROLLABLE = {"Canvas", "Text", "Listbox", "Treeview"}

        def _find_scrollable(widget):
            """Walk up the widget tree to find the first scrollable widget."""
            w = widget
            visited = set()
            while w is not None:
                try:
                    wid = str(w)
                    if wid in visited:
                        break
                    visited.add(wid)
                    if w.winfo_class() in SCROLLABLE:
                        return w
                    parent_path = w.winfo_parent()
                    if not parent_path:
                        break
                    w = w.nametowidget(parent_path)
                except Exception:
                    break
            return None

        def _scroll(widget, direction):
            """Scroll `widget` by `direction` units (+1 = down, -1 = up)."""
            scrollable = _find_scrollable(widget)
            if scrollable is None:
                return
            try:
                scrollable.yview_scroll(direction, "units")
            except Exception:
                pass

        def _on_mousewheel(event):
            try:
                w = self.winfo_containing(*self.winfo_pointerxy())
                if w is None:
                    return
            except Exception:
                return
            delta = getattr(event, "delta", 0)
            direction = -1 if delta > 0 else 1
            _scroll(w, direction)

        def _on_scroll_up(event):
            try:
                w = self.winfo_containing(*self.winfo_pointerxy())
                if w:
                    _scroll(w, -1)
            except Exception:
                pass

        def _on_scroll_down(event):
            try:
                w = self.winfo_containing(*self.winfo_pointerxy())
                if w:
                    _scroll(w, 1)
            except Exception:
                pass

        self.bind_all("<MouseWheel>", _on_mousewheel)
        self.bind_all("<Button-4>",   _on_scroll_up)    
        self.bind_all("<Button-5>",   _on_scroll_down)  

    # ── Sell Tab ──────────────────────────────────────────────────────────────
    def _build_sell_tab(self):
        c = self.colors
        f = self.tab_sell

        # ── Left: search + results list ───────────────────────────────────────
        left = ttk.Frame(f)
        left.pack(side="left", fill="both", expand=True, padx=10, pady=10)

        # Header
        tk.Label(left, text="◈  Sell Item Lookup",
                 font=("Georgia", 12, "bold"),
                 bg=c["bg"], fg=c["accent"]).pack(anchor="w", pady=(0, 8))

        # Search bar
        search_row = tk.Frame(left, bg=c["bg"])
        search_row.pack(fill="x", pady=(0, 6))
        tk.Label(search_row, text="Search:", bg=c["bg"], fg=c["fg"],
                 font=("Georgia", 10)).pack(side="left", padx=(0, 6))
        self.sell_entry = tk.Entry(search_row, textvariable=self.sell_search_var,
                                   width=36, bg=c["sel"], fg=c["fg"],
                                   insertbackground=c["fg"], relief="flat",
                                   font=("Georgia", 10))
        self.sell_entry.pack(side="left")
        self.sell_search_var.trace_add("write", self._on_sell_search)

        # Results listbox area
        tk.Label(left, text="Results  (click to select):",
                 bg=c["bg"], fg=c["fg"],
                 font=("Georgia", 8, "italic")).pack(anchor="w", pady=(4, 2))

        results_frame = tk.Frame(left, bg=c["bg"])
        results_frame.pack(fill="both", expand=True)

        self.sell_results_tree = ttk.Treeview(
            results_frame,
            columns=("name", "rarity", "type", "buy_price"),
            show="headings",
            selectmode="browse",
            height=18,
        )
        self.sell_results_tree.heading("name",      text="Name")
        self.sell_results_tree.heading("rarity",    text="Rarity")
        self.sell_results_tree.heading("type",      text="Type")
        self.sell_results_tree.heading("buy_price", text="Buy Price")
        self.sell_results_tree.column("name",      width=280, anchor="w")
        self.sell_results_tree.column("rarity",    width=90,  anchor="center")
        self.sell_results_tree.column("type",      width=180, anchor="w")
        self.sell_results_tree.column("buy_price", width=110, anchor="center")

        RARITY_COLORS = RARITY_COLORS_MAP
        for rarity, color in RARITY_COLORS.items():
            self.sell_results_tree.tag_configure(
                rarity.replace(" ", "_"), foreground=color)
        self.sell_results_tree.tag_configure("odd",  background=ROW_ODD)
        self.sell_results_tree.tag_configure("even", background=ROW_EVEN)

        vsb = ttk.Scrollbar(results_frame, orient="vertical",
                            command=self.sell_results_tree.yview)
        self.sell_results_tree.configure(yscrollcommand=vsb.set)
        self.sell_results_tree.pack(side="left", fill="both", expand=True)
        vsb.pack(side="right", fill="y")
        self.sell_results_tree.bind("<<TreeviewSelect>>", self._on_sell_result_select)

        right = tk.Frame(f, bg=c["hdr"], width=320)
        right.pack(side="right", fill="y", padx=(0, 10), pady=10)
        right.pack_propagate(False)

        tk.Label(right, text="Sell Price",
                 font=("Georgia", 12, "bold"),
                 bg=c["hdr"], fg=c["accent"]).pack(pady=(14, 4))
        ttk.Separator(right, orient="horizontal").pack(fill="x", padx=10)

        sell_canvas = tk.Canvas(right, bg=c["hdr"], highlightthickness=0)
        sell_vsb    = ttk.Scrollbar(right, orient="vertical",
                                    command=sell_canvas.yview)
        sell_canvas.configure(yscrollcommand=sell_vsb.set)
        sell_vsb.pack(side="right", fill="y")
        sell_canvas.pack(side="left", fill="both", expand=True)

        self.sell_panel = tk.Frame(sell_canvas, bg=c["hdr"])
        self._sell_panel_window = sell_canvas.create_window(
            (0, 0), window=self.sell_panel, anchor="nw")

        def _on_sell_panel_configure(event):
            sell_canvas.configure(scrollregion=sell_canvas.bbox("all"))
            sell_canvas.itemconfig(self._sell_panel_window,
                                   width=sell_canvas.winfo_width())

        self.sell_panel.bind("<Configure>", _on_sell_panel_configure)
        sell_canvas.bind("<Configure>",
                         lambda e: sell_canvas.itemconfig(
                             self._sell_panel_window, width=e.width))

        self._draw_sell_panel_empty()

    def _draw_sell_panel_empty(self):
        for w in self.sell_panel.winfo_children():
            w.destroy()
        tk.Label(self.sell_panel,
                 text="Search and select an item\nto calculate a sell price.",
                 bg=self.colors["hdr"], fg=self.colors["fg"],
                 font=("Georgia", 9, "italic"),
                 justify="center").pack(pady=30)

    def _draw_sell_panel(self, item: dict, buy_price: int):
        c = self.colors
        for w in self.sell_panel.winfo_children():
            w.destroy()

        rcolor = RARITY_COLORS_MAP.get(normalize_rarity(item.get("Rarity", "")), c["fg"])
        wrap   = 260

        # ── Name ──
        tk.Label(self.sell_panel, text=item.get("Name", ""),
                 bg=c["hdr"], fg=rcolor,
                 font=("Georgia", 12, "bold"),
                 wraplength=wrap, justify="left").pack(anchor="w", pady=(0, 4))
        ttk.Separator(self.sell_panel).pack(fill="x", pady=4)

        def field(label, val):
            if not val:
                return
            row = tk.Frame(self.sell_panel, bg=c["hdr"])
            row.pack(fill="x", pady=1)
            tk.Label(row, text=label, bg=c["hdr"], fg=c["accent"],
                     font=("Georgia", 9, "bold"), width=12,
                     anchor="w").pack(side="left")
            tk.Label(row, text=val, bg=c["hdr"], fg=c["fg"],
                     font=("Georgia", 10), anchor="w",
                     wraplength=wrap - 90, justify="left").pack(side="left")

        field("Rarity",     item.get("Rarity", ""))
        field("Type",       item.get("Type", ""))
        field("Source",     item.get("Source", ""))
        field("Attunement", item.get("Attunement", ""))
        field("Damage",     item.get("Damage", ""))
        field("Weight",     item.get("Weight", ""))
        field("List Price", item.get("Value", "") or "—")
        field("Buy Price",  format_currency(buy_price) if buy_price else item.get("Value", "") or "—")

        ttk.Separator(self.sell_panel).pack(fill="x", pady=10)

        # ── Sell % slider ──
        tk.Label(self.sell_panel, text="Shop's Buy Cut:",
                 bg=c["hdr"], fg=c["fg"],
                 font=("Georgia", 9, "bold")).pack(anchor="w", pady=(0, 4))

        slider_row = tk.Frame(self.sell_panel, bg=c["hdr"])
        slider_row.pack(fill="x", pady=(0, 4))

        self.sell_pct_var.set(80)
        ttk.Scale(slider_row, from_=10, to=100,
                  variable=self.sell_pct_var, orient="horizontal",
                  length=180, command=self._on_sell_slider).pack(side="left")

        self.sell_pct_disp = tk.Label(slider_row, text="80%",
                                       bg=c["hdr"], fg=c["accent"],
                                       font=("Georgia", 10, "bold"), width=5)
        self.sell_pct_disp.pack(side="left", padx=6)

        ttk.Separator(self.sell_panel).pack(fill="x", pady=6)

        # ── Offer price ──
        tk.Label(self.sell_panel, text="Offer to Seller:",
                 bg=c["hdr"], fg=c["fg"],
                 font=("Georgia", 9)).pack(anchor="w")
        self.sell_offer_disp = tk.Label(self.sell_panel, text="",
                                         bg=c["hdr"], fg="#1eff00",
                                         font=("Georgia", 16, "bold"))
        self.sell_offer_disp.pack(anchor="w", pady=(2, 0))

        self._update_sell_offer()

        # ── Description ──
        desc       = item.get("Text", "")
        table_data = item.get("Table", "")
        if desc or table_data:
            ttk.Separator(self.sell_panel).pack(fill="x", pady=(10, 6))
            tk.Label(self.sell_panel, text="Description",
                     bg=c["hdr"], fg=c["accent"],
                     font=("Georgia", 9, "bold"), anchor="w").pack(fill="x")
            if desc:
                prose = re.sub(r'(?<=[a-z])([.!?])(\)?)(?=[A-Z])', r'\1\2\n\n', desc)
                tk.Label(self.sell_panel, text=prose,
                         bg=c["sel"], fg=c["fg"],
                         font=("Georgia", 10),
                         wraplength=wrap, justify="left",
                         anchor="nw", padx=6, pady=6).pack(fill="x", pady=(4, 0))
            for tbl in parse_pipe_table(table_data):
                self._make_table_frame(
                    self.sell_panel, tbl["headers"], tbl["rows"]
                ).pack(fill="x", pady=(4, 0))

    def _on_sell_search(self, *_):
        q = self.sell_search_var.get().strip().lower()
        if not hasattr(self, "sell_results_tree"):
            return
        self.sell_results_tree.delete(*self.sell_results_tree.get_children())

        if len(q) < 2:
            return

        matches = [i for i in ALL_ITEMS_FLAT
                   if q in i.get("Name", "").lower()][:80]

        RARITY_COLORS = RARITY_COLORS_MAP
        for row_idx, item in enumerate(matches):
            calc_p_val = parse_given_cost(item.get("Value", ""))
            calc_p  = calc_p_val if calc_p_val else 0.0   # keep as float — sub-GP values matter
            rnorm   = normalize_rarity(item.get("Rarity", ""))
            r_tag   = rnorm.replace(" ", "_")
            parity  = "odd" if row_idx % 2 == 0 else "even"
            self.sell_results_tree.insert("", "end",
                values=(
                    item.get("Name", ""),
                    item.get("Rarity", "—"),
                    item.get("Type", "—"),
                    format_currency(calc_p),
                ),
                tags=(parity, r_tag),
                iid=f"sell_{row_idx}",
            )
            self.sell_results_tree.set(f"sell_{row_idx}", "buy_price",
                                        format_currency(calc_p))
            self._sell_result_data = getattr(self, "_sell_result_data", {})
            self._sell_result_data[f"sell_{row_idx}"] = (item, calc_p)

    def _on_sell_result_select(self, _=None):
        sel = self.sell_results_tree.selection()
        if not sel:
            return
        iid = sel[0]
        data = getattr(self, "_sell_result_data", {})
        if iid not in data:
            return
        item, buy_price = data[iid]
        self.sell_selected_item = {"item": item, "buy_price": buy_price}
        self._draw_sell_panel(item, buy_price)

    def _on_sell_slider(self, _=None):
        pct = int(float(self.sell_pct_var.get()))
        self.sell_pct_var.set(pct)
        if hasattr(self, "sell_pct_disp"):
            self.sell_pct_disp.configure(text=f"{pct}%")
        self._update_sell_offer()

    def _update_sell_offer(self):
        if not self.sell_selected_item:
            return
        pct   = int(self.sell_pct_var.get())
        buy_p = self.sell_selected_item["buy_price"]
        offer = max(0.01, float(buy_p) * pct / 100)
        if hasattr(self, "sell_offer_disp"):
            self.sell_offer_disp.configure(text=format_currency(offer))

    # ── Action Tab ────────────────────────────────────────────────────────────
    def _build_action_tab(self):
        c = self.colors
        f = self.tab_action

        left = ttk.Frame(f)
        left.pack(side="left", fill="both", expand=True)

        # Button bar
        btn_bar = tk.Frame(left, bg=c["hdr"], pady=6)
        btn_bar.pack(fill="x")
        ttk.Button(btn_bar, text="⚡ Generate Shop",
                   command=self._run_generate).pack(side="left", padx=6)
        ttk.Button(btn_bar, text="↻ Reroll (10-30%)",
                   command=self._reroll).pack(side="left", padx=6)
        ttk.Button(btn_bar, text="✖ Clear Shop",
                   style="Danger.TButton",
                   command=self._clear).pack(side="left", padx=6)
        ttk.Button(btn_bar, text="＋ Add Item",
                   command=self._open_add_item_dialog).pack(side="left", padx=6)

        # ── Discount / Markup slider ──
        tk.Frame(btn_bar, bg=c["sel"], width=2, height=26).pack(
            side="left", padx=(10, 8))

        tk.Label(btn_bar, text="Price Adjust:", bg=c["hdr"], fg=c["fg"],
                 font=("Georgia", 9)).pack(side="left")

        self.price_mod_slider = ttk.Scale(
            btn_bar, from_=50, to=125,
            variable=self.price_modifier,
            orient="horizontal", length=130,
            command=self._on_price_modifier)
        self.price_mod_slider.pack(side="left", padx=(4, 2))

        self.price_mod_label = tk.Label(
            btn_bar, text="100%", width=5,
            bg=c["hdr"], fg=c["accent"],
            font=("Georgia", 9, "bold"))
        self.price_mod_label.pack(side="left", padx=(0, 10))

        # Search bar
        tk.Label(btn_bar, text="Filter Items:", bg=c["hdr"], fg=c["fg"],
                 font=("Georgia", 9)).pack(side="right", padx=(0, 4))
        self.search_var = tk.StringVar()
        self.search_var.trace_add("write", lambda *_: self._populate_table(self.current_items))
        tk.Entry(btn_bar, textvariable=self.search_var, width=22,
                 bg=c["sel"], fg=c["fg"], insertbackground=c["fg"],
                 relief="flat", font=("Georgia", 9)).pack(side="right", padx=(0, 4))

        cols   = ("name", "rarity", "cost", "est_value", "quantity", "locked")
        hdrs   = ("Name", "Rarity", "Cost", "Est. Value", "Qty", "Locked")
        widths = (290, 100, 120, 110, 65, 60)

        tree_frame = ttk.Frame(left)
        tree_frame.pack(fill="both", expand=True, padx=6, pady=4)

        self.tree = ttk.Treeview(tree_frame, columns=cols,
                                  show="headings", selectmode="browse")

        for col, hdr, w in zip(cols, hdrs, widths):
            self.tree.heading(col, text=hdr,
                              command=lambda c=col: self._on_sort(c))
            self.tree.column(col, width=w,
                             anchor="w" if col == "name" else "center")

        self.tree.tag_configure("odd",          background=ROW_ODD)
        self.tree.tag_configure("even",         background=ROW_EVEN)
        self.tree.tag_configure("locked_odd",   background=ROW_LOCKED_ODD)
        self.tree.tag_configure("locked_even",  background=ROW_LOCKED_EVEN)
        self.tree.tag_configure("selected_row", background=ROW_SELECTED)

        RARITY_FG = RARITY_COLORS_MAP
        for rarity, color in RARITY_FG.items():
            tag = rarity.replace(" ", "_")
            self.tree.tag_configure(tag, foreground=color)

        vsb = ttk.Scrollbar(tree_frame, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=vsb.set)
        self.tree.pack(side="left", fill="both", expand=True)
        vsb.pack(side="right", fill="y")

        self.tree.bind("<<TreeviewSelect>>", self._on_select)
        self.tree.bind("<Double-1>",         self._on_double_click)

        self._update_display_columns()

        self.status_var = tk.StringVar(value="No shop generated.")
        tk.Label(left, textvariable=self.status_var,
                 bg=c["hdr"], fg=c["accent"],
                 font=("Georgia", 9), anchor="w").pack(fill="x", padx=6, pady=2)

        # ── Inspector panel ──────────────────────────────────────────────────

        self._action_left = left
        self._inspect_width_collapsed = 400
        self._inspect_width_expanded  = None

        self.inspect_panel = tk.Frame(f, bg=c["hdr"])
        self.inspect_panel.place(relx=1.0, rely=0.0,
                                  anchor="ne",
                                  width=self._inspect_width_collapsed,
                                  relheight=1.0)

        left.pack_configure(padx=(0, self._inspect_width_collapsed + 6))

        # ── Inspector header row (title + expand button) ──
        hdr_row = tk.Frame(self.inspect_panel, bg=c["hdr"])
        hdr_row.pack(fill="x", padx=8, pady=(10, 0))

        tk.Label(hdr_row, text="◈  Item Inspector",
                 font=("Georgia", 11, "bold"),
                 bg=c["hdr"], fg=c["accent"]).pack(side="left")

        self.expand_btn = tk.Label(
            hdr_row, text="⤢", font=("Georgia", 13),
            bg=c["hdr"], fg=c["fg"], cursor="hand2", padx=4)
        self.expand_btn.pack(side="right")
        self.expand_btn.bind("<Button-1>", lambda e: self._toggle_inspect_expand())
        self.expand_btn.bind("<Enter>",
            lambda e: self.expand_btn.configure(fg=c["accent"]))
        self.expand_btn.bind("<Leave>",
            lambda e: self.expand_btn.configure(fg=c["fg"]))

        ttk.Separator(self.inspect_panel, orient="horizontal").pack(
            fill="x", padx=8, pady=(4, 0))

        inspect_canvas = tk.Canvas(self.inspect_panel, bg=c["hdr"],
                                    highlightthickness=0)
        inspect_vsb = ttk.Scrollbar(self.inspect_panel, orient="vertical",
                                     command=inspect_canvas.yview)
        inspect_canvas.configure(yscrollcommand=inspect_vsb.set)
        inspect_vsb.pack(side="right", fill="y")
        inspect_canvas.pack(side="left", fill="both", expand=True)

        self.inspect_frame = tk.Frame(inspect_canvas, bg=c["hdr"])
        self._inspect_canvas_win = inspect_canvas.create_window(
            (0, 0), window=self.inspect_frame, anchor="nw")

        def _on_inspect_frame_configure(event):
            inspect_canvas.configure(scrollregion=inspect_canvas.bbox("all"))

        def _on_inspect_canvas_configure(event):
            inspect_canvas.itemconfig(
                self._inspect_canvas_win, width=event.width)

        self.inspect_frame.bind("<Configure>", _on_inspect_frame_configure)
        inspect_canvas.bind("<Configure>", _on_inspect_canvas_configure)

        self._clear_inspect()

    # ── Inspector expand / collapse ───────────────────────────────────────────
    def _toggle_inspect_expand(self):
        expanding = not self._inspect_expanded

        if self._inspect_width_expanded is None or expanding:
            self.update_idletasks()
            win_w = self.winfo_width()
            self._inspect_width_expanded = max(600, int(win_w * 0.56))

        w = (self._inspect_width_expanded if expanding
             else self._inspect_width_collapsed)

        self.inspect_panel.place_configure(width=w)
        self._action_left.pack_configure(padx=(0, w + 6))
        self._inspect_expanded = expanding
        self.expand_btn.configure(text="⤡" if expanding else "⤢")
        if self.selected_row:
            self._show_inspect(self.selected_row)

    # ── Sort ──────────────────────────────────────────────────────────────────
    def _on_sort(self, col: str):
        if self._sort_col == col:
            self._sort_asc = not self._sort_asc
        else:
            self._sort_col = col
            self._sort_asc = True
        self._populate_table(self.current_items)

    def _sorted_items(self, items: list[dict]) -> list[dict]:
        col = self._sort_col
        asc = self._sort_asc

        if not col or col == "rarity":
            return sorted(
                items,
                key=lambda i: (rarity_rank(i.get("rarity", "")),
                               i.get("name", "").lower()),
                reverse=not asc,
            )

        def key(item):
            if col == "name":
                return item.get("name", "").lower()
            if col == "cost":
                return parse_given_cost(item.get("cost_given", "")) or 0
            if col == "quantity":
                try: return int(item.get("quantity", "1") or "1")
                except: return 1
            if col == "locked":
                return int(item.get("locked", False))
            return str(item.get(col, "")).lower()

        return sorted(items, key=key, reverse=not asc)

    # ── Table ─────────────────────────────────────────────────────────────────
    def _populate_table(self, items: list[dict]):
        q = self.search_var.get().lower() if hasattr(self, "search_var") else ""
        self.tree.delete(*self.tree.get_children())

        visible = [i for i in items
                   if not q
                   or q in i["name"].lower()
                   or q in (i.get("rarity") or "").lower()
                   or q in (i.get("item_type") or "").lower()]

        visible = self._sorted_items(visible)

        for row_idx, item in enumerate(visible):
            is_locked  = item.get("locked", False)
            parity     = "odd" if row_idx % 2 == 0 else "even"
            bg_tag     = f"locked_{parity}" if is_locked else parity
            rarity_tag = normalize_rarity(item.get("rarity", "none")).replace(" ", "_")
            tags       = (bg_tag, rarity_tag)

            mod       = self.price_modifier.get()
            cost_disp = apply_price_mod(item.get("cost_given", ""), mod)
            qty_disp  = item.get("quantity", "1") or "1"
            lock_sym  = "◆" if is_locked else "◇"

            mkt_raw = item.get("market_price", "")
            if mkt_raw:
                try:
                    mkt_disp = f"{int(mkt_raw):,} gp"
                except ValueError:
                    mkt_disp = "—"
            else:
                mkt_disp = "—"

            self.tree.insert("", "end",
                values=(
                    item["name"],
                    item.get("rarity", ""),
                    cost_disp,
                    mkt_disp,
                    qty_disp,
                    lock_sym,
                ),
                tags=tags,
                iid=item["name"])

    # ── Inspector ─────────────────────────────────────────────────────────────
    def _clear_inspect(self):
        for w in self.inspect_frame.winfo_children():
            w.destroy()
        tk.Label(self.inspect_frame, text="Select an item to inspect.",
                 bg=self.colors["hdr"], fg=self.colors["fg"],
                 font=("Georgia", 9, "italic")).pack(pady=20)

    def _show_inspect(self, item: dict):
        for w in self.inspect_frame.winfo_children():
            w.destroy()
        if self._inspect_expanded:
            self._render_inspect_expanded(item)
        else:
            self._render_inspect_collapsed(item)

    def _make_table_frame(self, parent, headers: list, rows: list,
                          title: str = "") -> tk.Frame:

        c       = self.colors
        n_cols  = max(len(headers), 1)
        TTL_BG  = "#111128"
        HDR_BG  = "#1c1c3a"
        HDR_FG  = c["accent"]
        ROW_A   = "#1a1a30"
        ROW_B   = "#1e1e38"
        SEP_CLR = "#2a2a50"

        outer = tk.Frame(parent, bg=SEP_CLR, bd=0, relief="flat")

        if title:
            ttl_row = tk.Frame(outer, bg=TTL_BG)
            ttl_row.pack(fill="x", pady=(0, 1))
            tk.Label(
                ttl_row, text=title,
                bg=TTL_BG, fg=HDR_FG,
                font=("Georgia", 9, "bold italic"),
                padx=8, pady=4, anchor="w",
            ).pack(fill="x")

        hdr_row = tk.Frame(outer, bg=HDR_BG)
        hdr_row.pack(fill="x", pady=(0, 1))
        for col_idx, header in enumerate(headers):
            tk.Label(
                hdr_row, text=header,
                bg=HDR_BG, fg=HDR_FG,
                font=("Georgia", 9, "bold"),
                padx=8, pady=4, anchor="w",
            ).grid(row=0, column=col_idx, sticky="ew", padx=(0, 1))
            hdr_row.columnconfigure(col_idx, weight=1, minsize=60)

        for row_idx, row in enumerate(rows):
            bg = ROW_A if row_idx % 2 == 0 else ROW_B
            row_frame = tk.Frame(outer, bg=bg)
            row_frame.pack(fill="x", pady=(0, 1))
            for col_idx in range(n_cols):
                cell = row[col_idx] if col_idx < len(row) else ""
                tk.Label(
                    row_frame, text=cell,
                    bg=bg, fg=c["fg"],
                    font=("Georgia", 9),
                    padx=8, pady=3, anchor="nw",
                    wraplength=200, justify="left",
                ).grid(row=0, column=col_idx, sticky="nsew", padx=(0, 1))
                row_frame.columnconfigure(col_idx, weight=1, minsize=60)

        return outer

    def _render_description_rich(self, txt: tk.Text, raw_text: str,
                                  table_data: str = "") -> int:
        """Render description prose into txt widget, then any pipe-delimited tables.

        Returns an estimated line-height for auto-sizing the Text widget.
        """
        est_lines = 0

        # ── Prose ──────────────────────────────────────────────────────────────
        if raw_text:
            prose = re.sub(r'(?<=[a-z])([.!?])(\)?)(?=[A-Z])', r'\1\2\n\n', raw_text)
            txt.insert("end", prose)
            est_lines += max(1, len(prose) // 34) + prose.count("\n")

        # ── Tables from Table column ────────────────────────────────────────────
        tables = parse_pipe_table(table_data)
        for tbl in tables:
            if est_lines > 0:
                txt.insert("end", "\n\n")
                est_lines += 2
            frame = self._make_table_frame(txt, tbl["headers"], tbl["rows"])
            txt.window_create("end", window=frame, padx=2, pady=4)
            est_lines += len(tbl["rows"]) + 3

        return est_lines

    # ── Editable quantity widget ───────────────────────────────────────────────
    def _make_qty_editor(self, parent: tk.Widget, item: dict,
                         horizontal: bool = False):
        """Render a −/+ quantity editor for an item in the inspector panel.

        If *horizontal* is True, pack side-by-side (for expanded view row).
        Otherwise, pack vertically (for collapsed view field).
        When quantity reaches 0 the item is removed from the shop.
        """
        c = self.colors

        if not horizontal:
            tk.Label(parent, text="Quantity",
                     bg=c["hdr"], fg=c["accent"],
                     font=("Georgia", 9, "bold"), anchor="w").pack(fill="x", pady=(4, 0))

        qty_frame = tk.Frame(parent, bg=c["hdr"])
        if horizontal:
            qty_frame.pack(side="left")
        else:
            qty_frame.pack(anchor="w")

        try:
            current_qty = max(0, int(item.get("quantity", "1") or "1"))
        except (ValueError, TypeError):
            current_qty = 1

        qty_var = tk.IntVar(value=current_qty)
        qty_lbl = tk.Label(qty_frame, textvariable=qty_var, width=4,
                           bg=c["sel"], fg=c["fg"],
                           font=("Georgia", 11, "bold"),
                           anchor="center", relief="flat")

        def _change(delta: int):
            new_val = max(0, qty_var.get() + delta)
            qty_var.set(new_val)
            item["quantity"] = str(new_val)
            for shop_item in self.current_items:
                if shop_item["name"] == item["name"]:
                    shop_item["quantity"] = str(new_val)
                    break
            if new_val == 0:
                self.current_items = [i for i in self.current_items
                                      if i["name"] != item["name"]]
                self._populate_table(self.current_items)
                self.status_var.set(
                    f"✖  '{item['name']}' removed (quantity set to 0)")
                self._clear_inspect()
            else:
                self._populate_table(self.current_items)
                if item["name"] in self.tree.get_children():
                    self.tree.selection_set(item["name"])

        btn_minus = tk.Button(qty_frame, text="−", width=2,
                              bg=c["sel"], fg=c["fg"],
                              activebackground=c["hdr"], activeforeground=c["accent"],
                              relief="flat", font=("Georgia", 10, "bold"),
                              cursor="hand2", command=lambda: _change(-1))
        btn_plus  = tk.Button(qty_frame, text="＋", width=2,
                              bg=c["sel"], fg=c["fg"],
                              activebackground=c["hdr"], activeforeground=c["accent"],
                              relief="flat", font=("Georgia", 10, "bold"),
                              cursor="hand2", command=lambda: _change(1))

        btn_minus.pack(side="left", padx=(0, 2))
        qty_lbl.pack(side="left", padx=2)
        btn_plus.pack(side="left", padx=(2, 0))

    def _render_inspect_collapsed(self, item: dict):
        c = self.colors
        RARITY_FG = RARITY_COLORS_MAP
        rcolor = RARITY_FG.get(normalize_rarity(item.get("rarity", "")), c["fg"])
        wrap   = 370

        tk.Label(self.inspect_frame, text=item["name"],
                 bg=c["hdr"], fg=rcolor,
                 font=("Georgia", 12, "bold"),
                 wraplength=wrap, justify="left").pack(fill="x", pady=(0, 4))
        ttk.Separator(self.inspect_frame).pack(fill="x")

        if not item.get("_gallery"):
            btn = ttk.Button(self.inspect_frame, text="↻ Reroll This Item",
                             command=lambda i=item: self._reroll_single_item(i))
            btn.pack(anchor="w", pady=(6, 2))

        def field(label, val, multiline=False):
            if not val:
                return
            tk.Label(self.inspect_frame, text=label,
                     bg=c["hdr"], fg=c["accent"],
                     font=("Georgia", 9, "bold"), anchor="w").pack(fill="x", pady=(4, 0))
            if multiline:
                avg = 40
                wrapped = sum(max(1, -(-len(ln) // avg))
                              for ln in val.split("\n"))
                height = max(3, min(wrapped + val.count("\n"), 40))
                txt = tk.Text(self.inspect_frame, bg=c["sel"], fg=c["fg"],
                              wrap="word", height=height, relief="flat",
                              font=("Georgia", 10), padx=6, pady=4)
                txt.insert("1.0", val)
                txt.configure(state="disabled")
                txt.pack(fill="x")
            else:
                tk.Label(self.inspect_frame, text=val,
                         bg=c["hdr"], fg=c["fg"],
                         font=("Georgia", 10), anchor="w",
                         wraplength=wrap, justify="left").pack(fill="x")

        src = item.get("source", "")
        pg  = item.get("page", "")
        field("Item ID",    item.get("item_id"))
        field("Type",       item.get("item_type"))
        field("Rarity",     item.get("rarity"))
        field("Source",     f"{src} p.{pg}" if pg else src)
        field("Attunement", item.get("attunement"))
        field("Damage",     item.get("damage"))
        field("Properties", item.get("properties"))
        field("Mastery",    item.get("mastery"))
        field("Weight",     item.get("weight"))
        field("Tags",       item.get("tags"))

        field("Cost",      item.get("cost_given"))

        # ── Editable Quantity ──
        if not item.get("_gallery"):
            self._make_qty_editor(self.inspect_frame, item)

        # ── Market Price (random range estimate) ──
        mkt_raw = item.get("market_price", "")
        if mkt_raw:
            try:
                mkt_str = f"{int(mkt_raw):,} gp"
            except ValueError:
                mkt_str = ""
            if mkt_str:
                field("Est. Value", mkt_str)

        # ── Sane Magical Prices guide value ──
        sane_raw = item.get("sane_cost", "")
        if sane_raw:
            try:
                sane_str = f"{int(sane_raw):,} gp  (Sane Prices)"
            except ValueError:
                sane_str = ""
            if sane_str:
                field("Sane Cost", sane_str)

        if item.get("description") or item.get("table_data"):
            tk.Label(self.inspect_frame, text="Description",
                     bg=c["hdr"], fg=c["accent"],
                     font=("Georgia", 9, "bold"), anchor="w").pack(fill="x", pady=(4, 0))
            desc_text = item.get("description", "")
            table_data = item.get("table_data", "")
            if desc_text:
                prose = re.sub(r'(?<=[a-z])([.!?])(\)?)(?=[A-Z])', r'\1\2\n\n', desc_text)
                tk.Label(self.inspect_frame, text=prose,
                         bg=c["sel"], fg=c["fg"],
                         font=("Georgia", 10),
                         wraplength=wrap, justify="left",
                         anchor="nw", padx=6, pady=6).pack(fill="x")
            for tbl in parse_pipe_table(table_data):
                self._make_table_frame(
                    self.inspect_frame, tbl["headers"], tbl["rows"]
                ).pack(fill="x", pady=(4, 0))

    # ── Expanded layout (spacious, readable) ─────────────────────────────────
    def _render_inspect_expanded(self, item: dict):
        c   = self.colors
        pad = 16
        RARITY_FG = RARITY_COLORS_MAP
        rarity = item.get("rarity", "")
        rcolor = RARITY_FG.get(normalize_rarity(rarity), c["fg"])

        # ── Title ──
        title_frame = tk.Frame(self.inspect_frame, bg=c["hdr"])
        title_frame.pack(fill="x", padx=pad, pady=(12, 0))

        tk.Label(title_frame, text=item["name"],
                 bg=c["hdr"], fg=rcolor,
                 font=("Georgia", 17, "bold"),
                 wraplength=480, justify="left").pack(anchor="w")

        sub_parts = [p for p in [rarity.title(), item.get("item_type", "")] if p]
        if sub_parts:
            tk.Label(title_frame, text="  ·  ".join(sub_parts),
                     bg=c["hdr"], fg=c["fg"],
                     font=("Georgia", 10, "italic")).pack(anchor="w", pady=(3, 0))

        if not item.get("_gallery"):
            ttk.Button(title_frame, text="↻ Reroll This Item",
                       command=lambda i=item: self._reroll_single_item(i)
                       ).pack(anchor="w", pady=(8, 0))

        ttk.Separator(self.inspect_frame).pack(fill="x", padx=pad, pady=10)

        # ── Stats — single column, generous sizing ──
        src = item.get("source", "")
        pg  = item.get("page", "")

        stats = [
            ("Item ID",     item.get("item_id", "")),
            ("Type",        item.get("item_type", "")),
            ("Rarity",      rarity.title()),
            ("Source",      f"{src} p.{pg}" if pg else src),
            ("Attunement",  item.get("attunement", "")),
            ("Damage",      item.get("damage", "")),
            ("Properties",  item.get("properties", "")),
            ("Mastery",     item.get("mastery", "")),
            ("Weight",      item.get("weight", "")),
            ("Tags",        item.get("tags", "")),
            ("Cost",        item.get("cost_given", "")),
        ]

        mkt_raw = item.get("market_price", "")
        if mkt_raw:
            try:
                stats.append(("Est. Value", f"{int(mkt_raw):,} gp"))
            except ValueError:
                pass

        # Sane Magical Prices
        sane_raw = item.get("sane_cost", "")
        if sane_raw:
            try:
                stats.append(("Sane Cost", f"{int(sane_raw):,} gp")   )
            except ValueError:
                pass
        stats = [(lbl, val) for lbl, val in stats if val]

        stats_frame = tk.Frame(self.inspect_frame, bg=c["hdr"])
        stats_frame.pack(fill="x", padx=pad)

        for lbl, val in stats:
            row = tk.Frame(stats_frame, bg=c["hdr"])
            row.pack(fill="x", pady=4)
            tk.Label(row, text=lbl,
                     bg=c["hdr"], fg=c["accent"],
                     font=("Georgia", 9, "bold"),
                     width=11, anchor="w").pack(side="left", padx=(0, 10))
            tk.Label(row, text=val,
                     bg=c["hdr"], fg=c["fg"],
                     font=("Georgia", 10),
                     wraplength=380, justify="left",
                     anchor="w").pack(side="left", fill="x", expand=True)

        # ── Editable Quantity ──
        if not item.get("_gallery"):
            qty_row = tk.Frame(stats_frame, bg=c["hdr"])
            qty_row.pack(fill="x", pady=4)
            tk.Label(qty_row, text="Quantity",
                     bg=c["hdr"], fg=c["accent"],
                     font=("Georgia", 9, "bold"),
                     width=11, anchor="w").pack(side="left", padx=(0, 10))
            self._make_qty_editor(qty_row, item, horizontal=True)

        # ── Description ──
        desc       = item.get("description", "")
        table_data = item.get("table_data", "")
        if desc or table_data:
            ttk.Separator(self.inspect_frame).pack(
                fill="x", padx=pad, pady=(12, 8))

            tk.Label(self.inspect_frame, text="DESCRIPTION",
                     bg=c["hdr"], fg=c["accent"],
                     font=("Georgia", 10, "bold"),
                     anchor="w").pack(fill="x", padx=pad, pady=(0, 6))

            if desc:
                prose = re.sub(r'(?<=[a-z])([.!?])(\)?)(?=[A-Z])', r'\1\2\n\n', desc)
                desc_bg = "#16162a"
                desc_frame = tk.Frame(self.inspect_frame, bg=desc_bg,
                                      highlightbackground="#2a2a4a",
                                      highlightthickness=1)
                desc_frame.pack(fill="x", padx=pad, pady=(0, 8))
                tk.Label(desc_frame, text=prose,
                         bg=desc_bg, fg="#d8d0b8",
                         font=("Georgia", 11),
                         wraplength=max(300, self._inspect_width_expanded or 500) - pad * 2 - 28,
                         justify="left", anchor="nw",
                         padx=14, pady=12).pack(fill="x")

            for tbl in parse_pipe_table(table_data):
                self._make_table_frame(
                    self.inspect_frame, tbl["headers"], tbl["rows"]
                ).pack(fill="x", padx=pad, pady=(4, 0))

    # ── Settings Tab ──────────────────────────────────────────────────────────
    def _build_settings_tab(self):
        c = self.colors
        f = self.tab_settings

        outer = ttk.Frame(f)
        outer.pack(fill="both", expand=True, padx=20, pady=16)

        left_col = ttk.Frame(outer)
        left_col.pack(side="left", fill="y", padx=(0, 30))

        tk.Label(left_col, text="Shop Type",
                 font=("Georgia", 11, "bold"),
                 bg=c["bg"], fg=c["accent"]).pack(anchor="w", pady=(0, 6))
        shop_combo = ttk.Combobox(left_col, textvariable=self.current_shop_type,
                                  values=list(SHOP_TYPE_TO_POOL.keys()), width=22,
                                  state="readonly")
        shop_combo.pack(anchor="w", pady=(0, 4))
        shop_combo.bind("<<ComboboxSelected>>", self._on_shop_type_change)

        ttk.Separator(left_col, orient="horizontal").pack(fill="x", pady=10)

        tk.Label(left_col, text="City Size",
                 font=("Georgia", 11, "bold"),
                 bg=c["bg"], fg=c["accent"]).pack(anchor="w", pady=(0, 6))
        for size, (lo, hi) in CITY_SIZE_RANGES.items():
            tk.Radiobutton(left_col,
                           text=f"{size}  ({lo}–{hi} items)",
                           variable=self.city_size_var, value=size,
                           bg=c["bg"], fg=c["fg"], selectcolor=c["sel"],
                           activebackground=c["bg"], activeforeground=c["accent"],
                           font=("Georgia", 10)).pack(anchor="w", pady=2)

        ttk.Separator(left_col, orient="horizontal").pack(fill="x", pady=10)

        tk.Label(left_col, text="Wealth Level",
                 font=("Georgia", 11, "bold"),
                 bg=c["bg"], fg=c["accent"]).pack(anchor="w", pady=(0, 6))
        for wealth in WEALTH_DEFAULTS:
            tk.Radiobutton(left_col, text=wealth,
                           variable=self.wealth_var, value=wealth,
                           command=self._on_wealth_change,
                           bg=c["bg"], fg=c["fg"], selectcolor=c["sel"],
                           activebackground=c["bg"], activeforeground=c["accent"],
                           font=("Georgia", 10)).pack(anchor="w", pady=2)

        ttk.Separator(left_col, orient="horizontal").pack(fill="x", pady=10)

        tk.Label(left_col, text="Shop Mode",
                 font=("Georgia", 11, "bold"),
                 bg=c["bg"], fg=c["accent"]).pack(anchor="w", pady=(0, 6))

        mundane_chk = tk.Checkbutton(
            left_col,
            text="Mundane Only\n(common items at most)",
            variable=self.mundane_only_var,
            command=self._on_mundane_only_toggle,
            bg=c["bg"], fg=c["fg"],
            selectcolor=c["sel"],
            activebackground=c["bg"], activeforeground=c["accent"],
            font=("Georgia", 10),
            justify="left",
            anchor="w",
        )
        mundane_chk.pack(anchor="w", pady=2)
        tk.Label(left_col,
                 text="Overrides rarity sliders.\nGood for general stores\nand travelling merchants.",
                 bg=c["bg"], fg=c["fg"],
                 font=("Georgia", 8, "italic"),
                 justify="left").pack(anchor="w", pady=(0, 4))

        ttk.Separator(left_col, orient="horizontal").pack(fill="x", pady=10)

        # ── Table Display Settings ─────────────────────────────────────────────
        tk.Label(left_col, text="Table Display",
                 font=("Georgia", 11, "bold"),
                 bg=c["bg"], fg=c["accent"]).pack(anchor="w", pady=(0, 6))
        tk.Checkbutton(
            left_col, text="Show Quantity column",
            variable=self.show_qty_col,
            command=self._update_display_columns,
            bg=c["bg"], fg=c["fg"], selectcolor=c["sel"],
            activebackground=c["bg"], activeforeground=c["accent"],
            font=("Georgia", 10)).pack(anchor="w", pady=2)
        tk.Checkbutton(
            left_col, text="Show Est. Value column",
            variable=self.show_est_val_col,
            command=self._update_display_columns,
            bg=c["bg"], fg=c["fg"], selectcolor=c["sel"],
            activebackground=c["bg"], activeforeground=c["accent"],
            font=("Georgia", 10)).pack(anchor="w", pady=2)

        # Rarity sliders
        right_col = ttk.Frame(outer)
        right_col.pack(side="left", fill="y", padx=(0, 20))

        tk.Label(right_col, text="Rarity Distribution (%)",
                 font=("Georgia", 11, "bold"),
                 bg=c["bg"], fg=c["accent"]).pack(anchor="w", pady=(0, 4))
        tk.Label(right_col,
                 text="Adjust sliders to override wealth presets. Total should equal 100%.",
                 bg=c["bg"], fg=c["fg"],
                 font=("Georgia", 8, "italic")).pack(anchor="w", pady=(0, 8))

        RARITY_COLORS = RARITY_COLORS_MAP
        self.slider_labels: dict[str, tk.StringVar] = {}

        for rarity in ["common", "uncommon", "rare", "very rare", "legendary", "artifact"]:
            row = tk.Frame(right_col, bg=c["bg"])
            row.pack(fill="x", pady=4)
            color = RARITY_COLORS.get(rarity, c["fg"])
            tk.Label(row, text=rarity.title(), width=12, anchor="w",
                     bg=c["bg"], fg=color,
                     font=("Georgia", 10)).pack(side="left")

            lbl_var = tk.StringVar(value=f"{self.rarity_sliders[rarity].get():>3}%")
            self.slider_labels[rarity] = lbl_var

            ttk.Scale(row, from_=0, to=100,
                      variable=self.rarity_sliders[rarity],
                      orient="horizontal", length=260,
                      command=lambda v, r=rarity: self._on_slider(r, v)
                      ).pack(side="left", padx=8)

            tk.Label(row, textvariable=lbl_var, width=5,
                     bg=c["bg"], fg=color,
                     font=("Consolas", 10)).pack(side="left")

        self.total_pct_var = tk.StringVar(value="Total: 100%")
        self.total_pct_label = tk.Label(right_col, textvariable=self.total_pct_var,
                 bg=c["bg"], fg=c["accent"],
                 font=("Georgia", 10, "bold"))
        self.total_pct_label.pack(anchor="w", pady=(8, 4))

        ttk.Button(right_col, text="↺  Reset Distribution",
                   command=self._reset_distribution).pack(anchor="w")

        tag_col = ttk.Frame(outer)
        tag_col.pack(side="left", fill="both", expand=True)
        self._build_tag_filter(tag_col)

    # ── Table column visibility ────────────────────────────────────────────────
    def _update_display_columns(self, *_):
        """Show/hide Qty and Est. Value columns based on toggle settings."""
        if not hasattr(self, "tree"):
            return
        all_cols = ("name", "rarity", "cost", "est_value", "quantity", "locked")
        visible = [c for c in all_cols
                   if not (c == "est_value" and not self.show_est_val_col.get())
                   and not (c == "quantity"  and not self.show_qty_col.get())]
        self.tree["displaycolumns"] = visible

    # ── Add Item from Gallery dialog ──────────────────────────────────────────
    def _open_add_item_dialog(self):
        """Open a search dialog to add items from the item gallery to the shop."""
        c = self.colors
        dlg = tk.Toplevel(self)
        dlg.title("Add Item to Shop")
        dlg.geometry("720x520")
        dlg.configure(bg=c["bg"])
        dlg.transient(self)
        dlg.grab_set()

        # ── Search bar ──
        bar = tk.Frame(dlg, bg=c["hdr"], pady=6)
        bar.pack(fill="x")
        tk.Label(bar, text="Search:", bg=c["hdr"], fg=c["fg"],
                 font=("Georgia", 9)).pack(side="left", padx=(10, 4))
        search_var = tk.StringVar()
        tk.Entry(bar, textvariable=search_var, width=30,
                 bg=c["sel"], fg=c["fg"], insertbackground=c["fg"],
                 relief="flat", font=("Consolas", 9)).pack(side="left", padx=(0, 12))

        tk.Label(bar, text="Rarity:", bg=c["hdr"], fg=c["fg"],
                 font=("Georgia", 9)).pack(side="left")
        rarity_var = tk.StringVar(value="All")
        rarity_opts = ["All", "Mundane", "Common", "Uncommon", "Rare",
                       "Very Rare", "Legendary", "Artifact"]
        ttk.Combobox(bar, textvariable=rarity_var, values=rarity_opts,
                     width=12, state="readonly").pack(side="left", padx=(4, 12))

        result_lbl = tk.Label(bar, text="", bg=c["hdr"], fg=c["fg"],
                              font=("Georgia", 8, "italic"))
        result_lbl.pack(side="right", padx=10)

        # ── Treeview ──
        tree_frame = ttk.Frame(dlg)
        tree_frame.pack(fill="both", expand=True, padx=6, pady=4)

        add_cols   = ("name", "rarity", "type", "source", "value")
        add_hdrs   = ("Name", "Rarity", "Type", "Source", "Value")
        add_widths = (260, 95, 180, 80, 100)

        add_tree = ttk.Treeview(tree_frame, columns=add_cols,
                                show="headings", selectmode="browse")
        for col, hdr, w in zip(add_cols, add_hdrs, add_widths):
            add_tree.heading(col, text=hdr)
            add_tree.column(col, width=w,
                            anchor="w" if col in ("name", "type") else "center")

        RARITY_FG = RARITY_COLORS_MAP
        for rarity, color in RARITY_FG.items():
            add_tree.tag_configure(rarity.replace(" ", "_"), foreground=color)

        add_vsb = ttk.Scrollbar(tree_frame, orient="vertical", command=add_tree.yview)
        add_tree.configure(yscrollcommand=add_vsb.set)
        add_tree.pack(side="left", fill="both", expand=True)
        add_vsb.pack(side="right", fill="y")

        # ── Button bar ──
        btn_bar2 = tk.Frame(dlg, bg=c["hdr"], pady=6)
        btn_bar2.pack(fill="x")
        ttk.Button(btn_bar2, text="✕ Cancel",
                   command=dlg.destroy).pack(side="right", padx=10)
        add_btn = ttk.Button(btn_bar2, text="＋ Add to Shop",
                             command=lambda: self._add_item_from_dialog(add_tree, dlg))
        add_btn.pack(side="right", padx=(0, 6))
        tk.Label(btn_bar2, text="Double-click or select and press Add to Shop",
                 bg=c["hdr"], fg=c["fg"],
                 font=("Georgia", 8, "italic")).pack(side="left", padx=10)

        add_tree.bind("<Double-1>", lambda e: self._add_item_from_dialog(add_tree, dlg))

        # ── Populate / filter ──
        _all_rows = ALL_ITEMS_FLAT

        def _refresh(*_):
            q  = search_var.get().strip().lower()
            rf = rarity_var.get()
            add_tree.delete(*add_tree.get_children())
            shown = 0
            for row in _all_rows:
                r_norm = normalize_rarity(row.get("Rarity", ""))
                if rf != "All" and r_norm != rf.lower():
                    continue
                name = row.get("Name", "")
                typ  = row.get("Type", "")
                src  = row.get("Source", "")
                tags = row.get("Tags", "")
                if q and q not in name.lower() and q not in typ.lower() \
                       and q not in (r_norm) and q not in tags.lower():
                    continue
                r_tag = r_norm.replace(" ", "_")
                add_tree.insert("", "end", iid=name,
                                values=(name, row.get("Rarity", ""),
                                        typ, src, row.get("Value", "")),
                                tags=(r_tag,))
                shown += 1
                if shown >= 500:
                    break
            result_lbl.configure(
                text=f"{shown}{'+'if shown==500 else ''} result(s)")

        search_var.trace_add("write", _refresh)
        rarity_var.trace_add("write", _refresh)
        _refresh()

    def _add_item_from_dialog(self, tree: "ttk.Treeview", dlg: tk.Toplevel):
        """Add the selected item from the add-item dialog into the current shop."""
        sel = tree.selection()
        if not sel:
            return
        item_name = sel[0]

        raw = next((r for r in ALL_ITEMS_FLAT if r.get("Name") == item_name), None)
        if raw is None:
            return

        existing_names = {i["name"] for i in self.current_items}
        if item_name in existing_names:
            from tkinter import messagebox
            messagebox.showinfo("Already in Shop",
                                f'"{item_name}" is already in the shop.',
                                parent=dlg)
            return

        rarity_raw = raw.get("Rarity", "")
        mkt = generate_market_price(rarity_raw)
        quantity = str(generate_item_quantity(
            raw, self.city_size_var.get(), self.wealth_var.get()))

        new_item = {
            "item_id":     raw.get("Item ID", ""),
            "name":        item_name,
            "rarity":      rarity_raw,
            "item_type":   raw.get("Type", ""),
            "source":      raw.get("Source", ""),
            "page":        raw.get("Page", ""),
            "cost_given":  raw.get("Value", ""),
            "quantity":    quantity,
            "locked":      False,
            "attunement":  raw.get("Attunement", ""),
            "damage":      raw.get("Damage", ""),
            "properties":  raw.get("Properties", ""),
            "mastery":     raw.get("Mastery", ""),
            "weight":      raw.get("Weight", ""),
            "tags":        raw.get("Tags", ""),
            "description": raw.get("Text", ""),
            "table_data":  raw.get("Table", ""),
            "sane_cost":   raw.get("Sane_Cost", ""),
            "market_price": str(mkt) if mkt is not None else "",
        }
        self.current_items.append(new_item)
        self._populate_table(self.current_items)
        self.status_var.set(
            f"＋ Added '{item_name}' to shop  ({len(self.current_items)} items total)")
        dlg.destroy()

    # ── Shopkeeper Tab ────────────────────────────────────────────────────────
    def _build_shopkeeper_tab(self):
        c = self.colors
        f = self.tab_shopkeeper

        outer = ttk.Frame(f)
        outer.pack(fill="both", expand=True, padx=20, pady=16)

        tk.Label(outer, text="Shopkeeper Generator",
                 font=("Georgia", 14, "bold"),
                 bg=c["bg"], fg=c["accent"]).pack(anchor="w", pady=(0, 4))
        tk.Label(outer,
                 text="Generate a shopkeeper NPC for your current shop. "
                      "Fields save automatically when you save the shop.",
                 bg=c["bg"], fg=c["fg"],
                 font=("Georgia", 9, "italic")).pack(anchor="w", pady=(0, 12))

        ttk.Separator(outer, orient="horizontal").pack(fill="x", pady=(0, 12))

        content = tk.Frame(outer, bg=c["bg"])
        content.pack(anchor="nw", fill="x")

        sk_frame = tk.Frame(content, bg=c["bg"],
                            highlightbackground=c["sel"], highlightthickness=1)
        sk_frame.pack(fill="x", pady=(0, 4), ipadx=4)

        sk_row1 = tk.Frame(sk_frame, bg=c["bg"])
        sk_row1.pack(fill="x", padx=12, pady=(12, 6))

        name_col = tk.Frame(sk_row1, bg=c["bg"])
        name_col.pack(side="left", fill="x", expand=True, padx=(0, 12))
        tk.Label(name_col, text="Name:", bg=c["bg"], fg=c["accent"],
                 font=("Georgia", 9, "bold")).pack(anchor="w")
        tk.Entry(name_col, textvariable=self.shopkeeper_name_var,
                 bg=c["sel"], fg=c["fg"], insertbackground=c["fg"],
                 relief="flat", font=("Georgia", 10)).pack(fill="x", ipady=4)

        race_col = tk.Frame(sk_row1, bg=c["bg"])
        race_col.pack(side="left", fill="x", expand=True)
        tk.Label(race_col, text="Race:", bg=c["bg"], fg=c["accent"],
                 font=("Georgia", 9, "bold")).pack(anchor="w")
        self._sk_race_combo = ttk.Combobox(
            race_col, textvariable=self.shopkeeper_race_var,
            values=self._custom_races, width=18, font=("Georgia", 10))
        self._sk_race_combo.pack(fill="x", ipady=3)

        # Personality
        sk_row2 = tk.Frame(sk_frame, bg=c["bg"])
        sk_row2.pack(fill="x", padx=12, pady=(0, 6))
        tk.Label(sk_row2, text="Personality:", bg=c["bg"], fg=c["accent"],
                 font=("Georgia", 9, "bold")).pack(anchor="w")
        sk_pers_inner = tk.Frame(sk_row2, bg=c["bg"])
        sk_pers_inner.pack(fill="x")
        self._sk_personality_txt = tk.Text(
            sk_pers_inner, height=4,
            bg=c["sel"], fg=c["fg"], insertbackground=c["fg"],
            relief="flat", font=("Georgia", 10), wrap="word", padx=6, pady=4)
        sk_pers_vsb = ttk.Scrollbar(sk_pers_inner, orient="vertical",
                                     command=self._sk_personality_txt.yview)
        self._sk_personality_txt.configure(yscrollcommand=sk_pers_vsb.set)
        sk_pers_vsb.pack(side="right", fill="y")
        self._sk_personality_txt.pack(side="left", fill="x", expand=True)

        def _pers_to_var(e=None):
            self.shopkeeper_personality_var.set(
                self._sk_personality_txt.get("1.0", "end-1c"))
        def _pers_to_txt(*_):
            val = self.shopkeeper_personality_var.get()
            if val != self._sk_personality_txt.get("1.0", "end-1c"):
                self._sk_personality_txt.delete("1.0", "end")
                self._sk_personality_txt.insert("1.0", val)
        self._sk_personality_txt.bind("<KeyRelease>", _pers_to_var)
        self.shopkeeper_personality_var.trace_add("write", _pers_to_txt)

        # Appearance / Quirk
        sk_row3 = tk.Frame(sk_frame, bg=c["bg"])
        sk_row3.pack(fill="x", padx=12, pady=(0, 6))
        tk.Label(sk_row3, text="Appearance / Quirk:", bg=c["bg"], fg=c["accent"],
                 font=("Georgia", 9, "bold")).pack(anchor="w")
        sk_app_inner = tk.Frame(sk_row3, bg=c["bg"])
        sk_app_inner.pack(fill="x")
        self._sk_appearance_txt = tk.Text(
            sk_app_inner, height=4,
            bg=c["sel"], fg=c["fg"], insertbackground=c["fg"],
            relief="flat", font=("Georgia", 10), wrap="word", padx=6, pady=4)
        sk_app_vsb = ttk.Scrollbar(sk_app_inner, orient="vertical",
                                    command=self._sk_appearance_txt.yview)
        self._sk_appearance_txt.configure(yscrollcommand=sk_app_vsb.set)
        sk_app_vsb.pack(side="right", fill="y")
        self._sk_appearance_txt.pack(side="left", fill="x", expand=True)

        def _app_to_var(e=None):
            self.shopkeeper_appearance_var.set(
                self._sk_appearance_txt.get("1.0", "end-1c"))
        def _app_to_txt(*_):
            val = self.shopkeeper_appearance_var.get()
            if val != self._sk_appearance_txt.get("1.0", "end-1c"):
                self._sk_appearance_txt.delete("1.0", "end")
                self._sk_appearance_txt.insert("1.0", val)
        self._sk_appearance_txt.bind("<KeyRelease>", _app_to_var)
        self.shopkeeper_appearance_var.trace_add("write", _app_to_txt)

        # Buttons row
        sk_btn_row = tk.Frame(sk_frame, bg=c["bg"])
        sk_btn_row.pack(fill="x", padx=12, pady=(4, 12))
        ttk.Button(sk_btn_row, text="✦ Generate Shopkeeper",
                   command=self._generate_shopkeeper).pack(side="left")
        ttk.Button(sk_btn_row, text="✕ Clear",
                   command=self._clear_shopkeeper).pack(side="left", padx=(8, 0))
        tk.Label(sk_btn_row,
                 text="✔  Saves automatically with shop.",
                 bg=c["bg"], fg=c["fg"],
                 font=("Georgia", 8, "italic")).pack(side="left", padx=(16, 0))

    def _clear_shopkeeper(self):
        """Clear all shopkeeper fields."""
        self.shopkeeper_name_var.set("")
        self.shopkeeper_race_var.set("")
        self.shopkeeper_personality_var.set("")
        self.shopkeeper_appearance_var.set("")

    def _generate_shopkeeper(self):
        """Randomly fill shopkeeper fields using the pool data (respects custom race list)."""
        sk = generate_shopkeeper(self.current_shop_type.get())
        if self._custom_races:
            sk["race"] = random.choice(self._custom_races)
        self.shopkeeper_name_var.set(sk["name"])
        self.shopkeeper_race_var.set(sk["race"])
        self.shopkeeper_personality_var.set(sk["personality"])
        full_appearance = f"{sk['appearance']}. {sk['quirk']}"
        self.shopkeeper_appearance_var.set(full_appearance)

    # ── App Settings Window (gear icon) ───────────────────────────────────────
    def _open_app_settings_window(self):
        """Open (or raise) the global App Settings dialog."""
        if hasattr(self, "_app_settings_win") and self._app_settings_win and \
                self._app_settings_win.winfo_exists():
            self._app_settings_win.lift()
            self._app_settings_win.focus_force()
            return

        c = self.colors
        win = tk.Toplevel(self)
        win.title("App Settings")
        win.geometry("520x600")
        win.minsize(440, 480)
        win.configure(bg=c["hdr"])
        win.resizable(True, True)
        self._app_settings_win = win

        # ── Title ─────────────────────────────────────────────────────────────
        title_bar = tk.Frame(win, bg=c["hdr"])
        title_bar.pack(fill="x", padx=16, pady=(14, 0))
        tk.Label(title_bar, text="⚙  App Settings",
                 font=("Georgia", 13, "bold"),
                 bg=c["hdr"], fg=c["accent"]).pack(side="left")
        ttk.Separator(win, orient="horizontal").pack(fill="x", padx=16, pady=(10, 0))

        # ── Scrollable body ───────────────────────────────────────────────────
        body_canvas = tk.Canvas(win, bg=c["hdr"], highlightthickness=0)
        body_vsb    = ttk.Scrollbar(win, orient="vertical", command=body_canvas.yview)
        body_canvas.configure(yscrollcommand=body_vsb.set)
        body_vsb.pack(side="right", fill="y")
        body_canvas.pack(side="left", fill="both", expand=True)
        body = tk.Frame(body_canvas, bg=c["hdr"])
        _bwin = body_canvas.create_window((0, 0), window=body, anchor="nw")
        body.bind("<Configure>", lambda e: body_canvas.configure(
            scrollregion=body_canvas.bbox("all")))
        body_canvas.bind("<Configure>",
            lambda e: body_canvas.itemconfig(_bwin, width=e.width))

        pad = 16

        def section(text):
            tk.Label(body, text=text, font=("Georgia", 10, "bold"),
                     bg=c["hdr"], fg=c["accent"]).pack(anchor="w", padx=pad, pady=(14, 4))
            ttk.Separator(body, orient="horizontal").pack(fill="x", padx=pad, pady=(0, 6))

        def row_label(text):
            return tk.Label(body, text=text, bg=c["hdr"], fg=c["fg"],
                            font=("Georgia", 9), anchor="w")

        # ── Section: Defaults ─────────────────────────────────────────────────
        section("Startup Defaults")

        # Default shop type
        r = tk.Frame(body, bg=c["hdr"])
        r.pack(fill="x", padx=pad, pady=3)
        tk.Label(r, text="Default Shop Type:", bg=c["hdr"], fg=c["fg"],
                 font=("Georgia", 9), width=22, anchor="w").pack(side="left")
        _def_shop_var = tk.StringVar(value=self._app_settings["default_shop_type"])
        ttk.Combobox(r, textvariable=_def_shop_var,
                     values=list(SHOP_TYPE_TO_POOL.keys()),
                     width=18, state="readonly").pack(side="left")

        # Default city size
        r2 = tk.Frame(body, bg=c["hdr"])
        r2.pack(fill="x", padx=pad, pady=3)
        tk.Label(r2, text="Default City Size:", bg=c["hdr"], fg=c["fg"],
                 font=("Georgia", 9), width=22, anchor="w").pack(side="left")
        _def_city_var = tk.StringVar(value=self._app_settings["default_city_size"])
        ttk.Combobox(r2, textvariable=_def_city_var,
                     values=list(CITY_SIZE_RANGES.keys()),
                     width=18, state="readonly").pack(side="left")

        # Default wealth
        r3 = tk.Frame(body, bg=c["hdr"])
        r3.pack(fill="x", padx=pad, pady=3)
        tk.Label(r3, text="Default Wealth Level:", bg=c["hdr"], fg=c["fg"],
                 font=("Georgia", 9), width=22, anchor="w").pack(side="left")
        _def_wealth_var = tk.StringVar(value=self._app_settings["default_wealth"])
        ttk.Combobox(r3, textvariable=_def_wealth_var,
                     values=list(WEALTH_DEFAULTS.keys()),
                     width=18, state="readonly").pack(side="left")

        # Default price modifier
        r4 = tk.Frame(body, bg=c["hdr"])
        r4.pack(fill="x", padx=pad, pady=3)
        tk.Label(r4, text="Default Price Modifier %:", bg=c["hdr"], fg=c["fg"],
                 font=("Georgia", 9), width=22, anchor="w").pack(side="left")
        _def_price_var = tk.IntVar(value=self._app_settings["default_price_mod"])
        tk.Spinbox(r4, from_=10, to=500, textvariable=_def_price_var,
                   width=6, bg=c["sel"], fg=c["fg"],
                   buttonbackground=c["sel"], relief="flat",
                   font=("Georgia", 9)).pack(side="left")

        # ── Section: Behaviour ─────────────────────────────────────────────────
        section("Behaviour")

        _auto_name_var = tk.BooleanVar(value=self._app_settings["auto_name_on_change"])
        tk.Checkbutton(body, text="Auto-generate shop name when type changes",
                       variable=_auto_name_var,
                       bg=c["hdr"], fg=c["fg"], selectcolor=c["sel"],
                       activebackground=c["hdr"], activeforeground=c["accent"],
                       font=("Georgia", 9)).pack(anchor="w", padx=pad, pady=3)

        # Gallery per page
        r5 = tk.Frame(body, bg=c["hdr"])
        r5.pack(fill="x", padx=pad, pady=3)
        tk.Label(r5, text="Gallery items per page:", bg=c["hdr"], fg=c["fg"],
                 font=("Georgia", 9), width=24, anchor="w").pack(side="left")
        _gall_per_var = tk.IntVar(value=self._app_settings["gallery_per_page"])
        ttk.Combobox(r5, textvariable=_gall_per_var,
                     values=[50, 100, 150, 200, 250, 500],
                     width=6).pack(side="left")

        # ── Section: Race Pool Editor ─────────────────────────────────────────
        section("Shopkeeper Race Pool")
        tk.Label(body,
                 text="Races that can appear when randomly generating a shopkeeper.\n"
                      "Edit freely — one race per line.",
                 bg=c["hdr"], fg=c["fg"],
                 font=("Georgia", 8, "italic"),
                 justify="left").pack(anchor="w", padx=pad, pady=(0, 6))

        race_frame = tk.Frame(body, bg=c["hdr"])
        race_frame.pack(fill="x", padx=pad, pady=(0, 6))

        race_txt = tk.Text(race_frame, height=10, width=30,
                           bg=c["sel"], fg=c["fg"], insertbackground=c["fg"],
                           relief="flat", font=("Consolas", 9),
                           wrap="none", padx=6, pady=4)
        race_vsb = ttk.Scrollbar(race_frame, orient="vertical", command=race_txt.yview)
        race_txt.configure(yscrollcommand=race_vsb.set)
        race_vsb.pack(side="right", fill="y")
        race_txt.pack(side="left", fill="both", expand=True)

        # Populate with current custom races
        race_txt.insert("1.0", "\n".join(self._custom_races))

        race_btn_row = tk.Frame(body, bg=c["hdr"])
        race_btn_row.pack(fill="x", padx=pad, pady=(0, 4))

        def _reset_races():
            race_txt.delete("1.0", "end")
            race_txt.insert("1.0", "\n".join(SHOPKEEPER_POOLS["races"]))

        ttk.Button(race_btn_row, text="↺ Reset to Defaults",
                   command=_reset_races).pack(side="left")
        tk.Label(race_btn_row,
                 text="Changes apply when you click Save.",
                 bg=c["hdr"], fg=c["fg"],
                 font=("Georgia", 7, "italic")).pack(side="left", padx=8)

        # ── Section: Culture Tags ──────────────────────────────────────────────
        section("About")
        tk.Label(body,
                 text="Settings apply immediately on Save.\n"
                      "Race pool changes update the shopkeeper\n"
                      "dropdown in Stock Settings.",
                 bg=c["hdr"], fg=c["fg"],
                 font=("Georgia", 8, "italic"),
                 justify="left").pack(anchor="w", padx=pad, pady=(0, 10))

        # ── Footer: Save / Cancel ─────────────────────────────────────────────
        ttk.Separator(win, orient="horizontal").pack(fill="x", padx=16, pady=(4, 0))
        footer = tk.Frame(win, bg=c["hdr"])
        footer.pack(fill="x", padx=16, pady=(6, 10))

        def _save_app_settings():
            raw_races = [ln.strip() for ln in race_txt.get("1.0", "end").splitlines()
                         if ln.strip()]
            self._custom_races = raw_races if raw_races else list(SHOPKEEPER_POOLS["races"])
            if hasattr(self, "_sk_race_combo") and self._sk_race_combo.winfo_exists():
                self._sk_race_combo.configure(values=self._custom_races)

            self._app_settings["default_shop_type"]   = _def_shop_var.get()
            self._app_settings["default_city_size"]   = _def_city_var.get()
            self._app_settings["default_wealth"]      = _def_wealth_var.get()
            self._app_settings["default_price_mod"]   = _def_price_var.get()
            self._app_settings["auto_name_on_change"] = _auto_name_var.get()
            self._app_settings["gallery_per_page"]    = int(_gall_per_var.get())

            self._gallery_per_page.set(self._app_settings["gallery_per_page"])

            win.destroy()

        ttk.Button(footer, text="✔  Save", command=_save_app_settings).pack(side="right")
        ttk.Button(footer, text="✕  Cancel", command=win.destroy).pack(side="right", padx=(0, 6))

    # ── Tag filter UI ────────────────────────────────────────────────────────
    def _build_tag_filter(self, parent: tk.Frame):
        """Build the collapsible tag category sections inside parent."""
        c = self.colors

        hdr = tk.Frame(parent, bg=c["bg"])
        hdr.pack(fill="x", pady=(0, 8))
        tk.Label(hdr, text="Tag Filters",
                 font=("Georgia", 11, "bold"),
                 bg=c["bg"], fg=c["accent"]).pack(side="left")
        tk.Label(hdr,
                 text="Click once to include (✓), again to exclude (✗), again to clear.",
                 bg=c["bg"], fg=c["fg"],
                 font=("Georgia", 8, "italic")).pack(side="left", padx=(10, 0))
        ttk.Button(hdr, text="✕ Clear All",
                   command=self._clear_tag_filters).pack(side="right")
        ttk.Button(hdr, text="☑ Include All",
                   command=self._select_all_tag_filters).pack(side="right", padx=(0, 4))

        self.tag_active_label = tk.Label(hdr, text="",
                                          bg=c["bg"], fg="#ff9900",
                                          font=("Georgia", 8, "bold"))
        self.tag_active_label.pack(side="right", padx=6)

        # Scrollable canvas for all category sections
        canvas_frame = tk.Frame(parent, bg=c["bg"])
        canvas_frame.pack(fill="both", expand=True)

        canvas = tk.Canvas(canvas_frame, bg=c["bg"], highlightthickness=0)
        vsb = ttk.Scrollbar(canvas_frame, orient="vertical", command=canvas.yview)
        canvas.configure(yscrollcommand=vsb.set)
        vsb.pack(side="right", fill="y")
        canvas.pack(side="left", fill="both", expand=True)

        inner = tk.Frame(canvas, bg=c["bg"])
        win = canvas.create_window((0, 0), window=inner, anchor="nw")

        def _on_inner_configure(e):
            canvas.configure(scrollregion=canvas.bbox("all"))
        def _on_canvas_configure(e):
            canvas.itemconfig(win, width=e.width)
        inner.bind("<Configure>", _on_inner_configure)
        canvas.bind("<Configure>", _on_canvas_configure)

        # Build one collapsible section per category
        RARITY_COLORS = RARITY_COLORS_MAP
        for cat_name, tags in TAG_CATEGORIES.items():
            self._build_tag_section(inner, cat_name, tags, c, RARITY_COLORS)

    def _build_tag_section(self, parent, cat_name: str, tags: list[str],
                            c: dict, rarity_colors: dict):
        """Build one collapsible category section with 3-state cycle buttons."""
        section = tk.Frame(parent, bg=c["bg"],
                           highlightbackground=c["sel"],
                           highlightthickness=1)
        section.pack(fill="x", padx=4, pady=3)

        collapsed = tk.BooleanVar(value=True)

        hdr = tk.Frame(section, bg=c["sel"], cursor="hand2")
        hdr.pack(fill="x")

        arrow_lbl = tk.Label(hdr, text="▶", font=("Consolas", 8),
                              bg=c["sel"], fg=c["accent"], width=2)
        arrow_lbl.pack(side="left", padx=(6, 2))
        tk.Label(hdr, text=cat_name, font=("Georgia", 9, "bold"),
                 bg=c["sel"], fg=c["fg"]).pack(side="left", pady=4)

        count_var = tk.StringVar(value="")
        count_lbl = tk.Label(hdr, textvariable=count_var,
                              bg=c["sel"], fg="#ff9900",
                              font=("Consolas", 8))
        count_lbl.pack(side="right", padx=8)

        body = tk.Frame(section, bg=c["bg"])

        # ── State colours ──────────────────────────────────────────────────────
        STATE_FG  = {0: c["fg"],    1: "#1eff00", 2: "#ff4444"}
        STATE_BG  = {0: c["bg"],    1: "#0d1f0d", 2: "#1f0d0d"}
        STATE_PFX = {0: "  ",       1: "✓ ",      2: "✗ "}

        def _refresh_count():
            n_inc = sum(1 for t in tags
                        if self._tag_state_vars.get(t, tk.IntVar()).get() == 1)
            n_exc = sum(1 for t in tags
                        if self._tag_state_vars.get(t, tk.IntVar()).get() == 2)
            parts = []
            if n_inc: parts.append(f"{n_inc} incl")
            if n_exc: parts.append(f"{n_exc} excl")
            count_var.set(" / ".join(parts))

        def _toggle(_=None):
            if collapsed.get():
                body.pack(fill="x", padx=8, pady=(4, 6))
                arrow_lbl.configure(text="▼")
                collapsed.set(False)
            else:
                body.pack_forget()
                arrow_lbl.configure(text="▶")
                collapsed.set(True)

        hdr.bind("<Button-1>", _toggle)
        for child in hdr.winfo_children():
            child.bind("<Button-1>", _toggle)

        cols = 4
        for idx, tag in enumerate(tags):
            var = tk.IntVar(value=0)
            self._tag_state_vars[tag] = var
            btn_ref: list = []   # mutable cell for the button reference

            def _cycle(t=tag, v=var, br=btn_ref, rf=_refresh_count):
                new_state = (v.get() + 1) % 3
                v.set(new_state)
                # Update include / exclude sets
                self.active_tag_filters.discard(t)
                self.excluded_tag_filters.discard(t)
                if new_state == 1:
                    self.active_tag_filters.add(t)
                elif new_state == 2:
                    self.excluded_tag_filters.add(t)
                # Repaint the button
                if br:
                    br[0].configure(
                        text=STATE_PFX[new_state] + t,
                        fg=STATE_FG[new_state],
                        bg=STATE_BG[new_state],
                    )
                rf()
                self._update_tag_summary_label()

            btn = tk.Button(
                body,
                text=STATE_PFX[0] + tag,
                command=_cycle,
                fg=STATE_FG[0], bg=STATE_BG[0],
                activeforeground=c["accent"],
                activebackground=c["sel"],
                relief="flat", bd=0,
                font=("Georgia", 8),
                anchor="w", padx=2,
            )
            btn_ref.append(btn)
            btn.grid(row=idx // cols, column=idx % cols, sticky="w", padx=2, pady=1)

    def _update_tag_summary_label(self):
        """Refresh the global 'N incl / N excl' label above the tag panels."""
        if not hasattr(self, "tag_active_label"):
            return
        n_inc = len(self.active_tag_filters)
        n_exc = len(self.excluded_tag_filters)
        parts = []
        if n_inc: parts.append(f"{n_inc} included")
        if n_exc: parts.append(f"{n_exc} excluded")
        self.tag_active_label.configure(text=" / ".join(parts))

    def _clear_tag_filters(self):
        self.active_tag_filters.clear()
        self.excluded_tag_filters.clear()
        c = self.colors
        for tag, var in self._tag_state_vars.items():
            var.set(0)
        # Repaint all buttons back to neutral — walk every tag section body
        self._repaint_all_tag_buttons()
        if hasattr(self, "tag_active_label"):
            self.tag_active_label.configure(text="")

    def _select_all_tag_filters(self):
        """Set every tag to include state."""
        self.active_tag_filters.clear()
        self.excluded_tag_filters.clear()
        c = self.colors
        for tag, var in self._tag_state_vars.items():
            var.set(1)
            self.active_tag_filters.add(tag)
        self._repaint_all_tag_buttons()
        self._update_tag_summary_label()

    def _repaint_all_tag_buttons(self):
        """Walk the widget tree and repaint any tag cycle-buttons to match state."""
        STATE_FG  = {0: self.colors["fg"], 1: "#1eff00", 2: "#ff4444"}
        STATE_BG  = {0: self.colors["bg"], 1: "#0d1f0d", 2: "#1f0d0d"}
        STATE_PFX = {0: "  ",             1: "✓ ",      2: "✗ "}
        for tag, var in self._tag_state_vars.items():
            s = var.get()
            # Buttons store their tag name inside the text — find by matching
            for widget in self._iter_tag_buttons():
                txt = widget.cget("text")
                # strip prefix (2 chars) to get bare tag name
                if len(txt) >= 2 and txt[2:] == tag:
                    widget.configure(
                        text=STATE_PFX[s] + tag,
                        fg=STATE_FG[s],
                        bg=STATE_BG[s],
                    )
                    break

    def _iter_tag_buttons(self):
        """Yield all tk.Button widgets that live inside tag section bodies."""
        def _recurse(w):
            if isinstance(w, tk.Button):
                yield w
            for child in w.winfo_children():
                yield from _recurse(child)
        if hasattr(self, "tab_settings"):
            yield from _recurse(self.tab_settings)

    def _on_slider(self, rarity: str, value: str):
        """Move one slider; if total would exceed 100%, clamp it and reduce
        other sliders proportionally to keep the sum at exactly 100."""
        new_val = int(float(value))
        self.rarity_sliders[rarity].set(new_val)

        others   = [r for r in self.rarity_sliders if r != rarity]
        others_sum = sum(self.rarity_sliders[r].get() for r in others)
        total      = new_val + others_sum

        if total > 100:
            excess = total - 100
            # Distribute the excess reduction across the other sliders,
            # proportionally — but never drop one below 0.
            reducible = [(r, self.rarity_sliders[r].get()) for r in others
                         if self.rarity_sliders[r].get() > 0]
            reducible_sum = sum(v for _, v in reducible)

            if reducible_sum > 0:
                for r, v in reducible:
                    cut = min(v, round(excess * v / reducible_sum))
                    self.rarity_sliders[r].set(max(0, v - cut))
                # Fix any rounding leftover by adjusting the largest reducible
                remaining = sum(self.rarity_sliders[r].get()
                                for r in others) + new_val - 100
                if remaining > 0:
                    for r, v in sorted(reducible, key=lambda x: -x[1]):
                        cur = self.rarity_sliders[r].get()
                        if cur > 0:
                            self.rarity_sliders[r].set(max(0, cur - remaining))
                            break
            else:
                # No other slider has room — clamp this one
                self.rarity_sliders[rarity].set(100 - others_sum)

        # Refresh all labels + total
        for r, var in self.rarity_sliders.items():
            self.slider_labels[r].set(f"{var.get():>3}%")
        total = sum(v.get() for v in self.rarity_sliders.values())
        color = self.colors["accent"] if total == 100 else "#ff4444"
        self.total_pct_var.set(f"Total: {total}%")

    def _on_wealth_change(self):
        wealth   = self.wealth_var.get()
        defaults = WEALTH_DEFAULTS.get(wealth, {})
        for rarity, var in self.rarity_sliders.items():
            val = defaults.get(rarity, 0)
            var.set(val)
            self.slider_labels[rarity].set(f"{val:>3}%")
        total = sum(v.get() for v in self.rarity_sliders.values())
        color = self.colors["accent"] if total == 100 else "#ff4444"
        self.total_pct_var.set(f"Total: {total}%")

    def _reset_distribution(self):
        """Reset sliders to the currently selected wealth preset."""
        self._on_wealth_change()

    # ── Price modifier ────────────────────────────────────────────────────────
    def _on_price_modifier(self, _=None):
        mod = int(float(self.price_modifier.get()))
        self.price_modifier.set(mod)
        self.price_mod_label.configure(text=f"{mod}%")
        # Highlight label when not at 100%
        color = "#ff9900" if mod != 100 else self.colors["accent"]
        self.price_mod_label.configure(fg=color)
        self._populate_table(self.current_items)

    # ── Save Tab ──────────────────────────────────────────────────────────────
    def _build_save_tab(self):
        c = self.colors
        f = self.tab_save

        left = ttk.Frame(f)
        left.pack(side="left", fill="both", expand=True, padx=8, pady=8)

        tk.Label(left, text="Saved Campaigns",
                 font=("Georgia", 11, "bold"),
                 bg=c["bg"], fg=c["accent"]).pack(anchor="w")

        tree_f = ttk.Frame(left)
        tree_f.pack(fill="both", expand=True, pady=4)

        self.save_tree = ttk.Treeview(tree_f, show="tree headings",
                                       columns=("info",), selectmode="browse")
        self.save_tree.heading("#0",   text="Campaign / Town / Shop")
        self.save_tree.heading("info", text="Details")
        self.save_tree.column("#0",   width=240)
        self.save_tree.column("info", width=200)
        vsb2 = ttk.Scrollbar(tree_f, orient="vertical", command=self.save_tree.yview)
        self.save_tree.configure(yscrollcommand=vsb2.set)
        self.save_tree.pack(side="left", fill="both", expand=True)
        vsb2.pack(side="right", fill="y")

        btn_f = tk.Frame(left, bg=c["bg"])
        btn_f.pack(fill="x", pady=4)
        ttk.Button(btn_f, text="Load Shop",
                   command=self._load_selected_shop).pack(side="left", padx=4)
        ttk.Button(btn_f, text="✖ Delete",
                   style="Danger.TButton",
                   command=self._delete_selected).pack(side="left", padx=4)
        ttk.Button(btn_f, text="Export JSON",
                   command=self._export_json).pack(side="left", padx=4)
        ttk.Button(btn_f, text="Import JSON",
                   command=self._import_json).pack(side="left", padx=4)

        # Save form
        right = ttk.Frame(f)
        right.pack(side="right", fill="y", padx=8, pady=8)

        tk.Label(right, text="Save Current Shop",
                 font=("Georgia", 11, "bold"),
                 bg=c["bg"], fg=c["accent"]).pack(anchor="w", pady=(0, 8))

        self.save_campaign_var = tk.StringVar()
        self.save_town_var     = tk.StringVar()

        for label, var in [("Campaign Name:", self.save_campaign_var),
                            ("Town/Location:", self.save_town_var)]:
            tk.Label(right, text=label, bg=c["bg"], fg=c["fg"],
                     font=("Georgia", 9)).pack(anchor="w")
            tk.Entry(right, textvariable=var, width=30,
                     bg=c["sel"], fg=c["fg"],
                     insertbackground=c["fg"], relief="flat").pack(anchor="w", pady=(0, 8))

        ttk.Button(right, text="Save Shop",
                   command=self._save_shop).pack(anchor="w", pady=4)
        ttk.Separator(right, orient="horizontal").pack(fill="x", pady=8)

        tk.Label(right, text="Shop Notes:",
                 bg=c["bg"], fg=c["fg"],
                 font=("Georgia", 9)).pack(anchor="w")
        notes_frame = tk.Frame(right, bg=c["bg"])
        notes_frame.pack(fill="both", expand=True, pady=(0, 8))
        self.shop_notes_widget = tk.Text(
            notes_frame, width=30, height=7,
            bg=c["sel"], fg=c["fg"],
            insertbackground=c["fg"],
            relief="flat", font=("Georgia", 9),
            wrap="word",
        )
        notes_vsb = ttk.Scrollbar(notes_frame, orient="vertical",
                                   command=self.shop_notes_widget.yview)
        self.shop_notes_widget.configure(yscrollcommand=notes_vsb.set)
        notes_vsb.pack(side="right", fill="y")
        self.shop_notes_widget.pack(side="left", fill="both", expand=True)

        ttk.Separator(right, orient="horizontal").pack(fill="x", pady=(0, 8))

        self.save_status_var = tk.StringVar(value="")
        tk.Label(right, textvariable=self.save_status_var,
                 bg=c["bg"], fg=c["accent"],
                 font=("Georgia", 9, "italic"),
                 wraplength=220).pack(anchor="w")

    # ── Core actions ──────────────────────────────────────────────────────────
    def _on_mundane_only_toggle(self):
        """Grey out / restore rarity sliders when mundane-only mode is toggled."""
        state = "disabled" if self.mundane_only_var.get() else "normal"
        if hasattr(self, "_rarity_slider_widgets"):
            for widget in self._rarity_slider_widgets:
                try:
                    widget.configure(state=state)
                except Exception:
                    pass

    def _get_rarity_weights(self) -> dict[str, int]:
        if self.mundane_only_var.get():
            # Force mundane + common only; cap at common
            return {"common": 60, "uncommon": 0, "rare": 0,
                    "very rare": 0, "legendary": 0, "artifact": 0}
        return {r: v.get() for r, v in self.rarity_sliders.items()}

    def _get_item_count(self) -> int:
        lo, hi = CITY_SIZE_RANGES.get(self.city_size_var.get(), (15, 25))
        return random.randint(lo, hi)

    def _reroll_single_item(self, item: dict):
        """Replace one unlocked item in the shop with a fresh pick of the same rarity."""
        if item.get("locked"):
            return
        shop_type = self.current_shop_type.get()
        if shop_type not in ALL_ITEMS or not ALL_ITEMS[shop_type]:
            return

        target_rarity = normalize_rarity(item.get("rarity", ""))
        existing_names = {i["name"] for i in self.current_items if i["name"] != item["name"]}
        excl    = self.excluded_tag_filters or set()
        incl    = self.active_tag_filters   or set()

        def _pool_filter(x: dict) -> bool:
            if x["Name"] in existing_names:
                return False
            item_tags = {t.strip() for t in x.get("Tags", "").split(",") if t.strip()}
            if excl and (item_tags & excl):
                return False
            if incl and not (item_tags & incl):
                return False
            if self.mundane_only_var.get():
                r = normalize_rarity(x.get("Rarity", ""))
                if r not in ("mundane", "none", "common", ""):
                    return False
            return True

        # Build a pool of same-rarity candidates not already in the shop
        pool = [x for x in ALL_ITEMS[shop_type]
                if normalize_rarity(x.get("Rarity", "")) == target_rarity
                and _pool_filter(x)]
        # Fallback: any rarity if pool is empty
        if not pool:
            pool = [x for x in ALL_ITEMS[shop_type] if _pool_filter(x)]
        if not pool:
            return

        chosen = random.choice(pool)
        rarity_raw = chosen.get("Rarity", "")
        mkt = generate_market_price(rarity_raw)
        new_item = {
            "item_id":      chosen.get("Item ID", ""),
            "name":         chosen.get("Name", ""),
            "rarity":       rarity_raw,
            "item_type":    chosen.get("Type", ""),
            "source":       chosen.get("Source", ""),
            "page":         chosen.get("Page", ""),
            "cost_given":   chosen.get("Value", ""),
            "quantity":     str(generate_item_quantity(
                                chosen,
                                self.city_size_var.get(),
                                self.wealth_var.get())),
            "locked":       False,
            "attunement":   chosen.get("Attunement", ""),
            "damage":       chosen.get("Damage", ""),
            "properties":   chosen.get("Properties", ""),
            "mastery":      chosen.get("Mastery", ""),
            "weight":       chosen.get("Weight", ""),
            "tags":         chosen.get("Tags", ""),
            "description":  chosen.get("Text", ""),
            "table_data":   chosen.get("Table", ""),
            "sane_cost":    chosen.get("Sane_Cost", ""),
            "market_price": str(mkt) if mkt is not None else "",
        }

        # Swap in place to preserve list order
        for idx, i in enumerate(self.current_items):
            if i["name"] == item["name"]:
                self.current_items[idx] = new_item
                break

        self._populate_table(self.current_items)
        self.selected_row = new_item
        self._show_inspect(new_item)
        # Reselect the new row in the tree
        try:
            self.tree.selection_set(new_item["name"])
            self.tree.see(new_item["name"])
        except Exception:
            pass
        self.status_var.set(f"↻  Rerolled '{item['name']}' → '{new_item['name']}'")

    def _run_generate(self):
        shop_type = self.current_shop_type.get()
        if not shop_type:
            messagebox.showerror("Error", "Please select a shop type.")
            return
        count   = self._get_item_count()
        weights = self._get_rarity_weights()
        if sum(weights.values()) == 0:
            messagebox.showwarning("Warning", "All weights are 0 — using Average defaults.")
            weights = WEALTH_DEFAULTS["Average"]
        self.current_items = generate_shop_items(
            shop_type, count, weights,
            tag_filters=self.active_tag_filters   if self.active_tag_filters   else None,
            tag_excludes=self.excluded_tag_filters if self.excluded_tag_filters else None,
            city_size=self.city_size_var.get(),
            wealth=self.wealth_var.get(),
            mundane_only=self.mundane_only_var.get())
        self._populate_table(self.current_items)
        self.status_var.set(
            f"✔  Generated {len(self.current_items)} items for {shop_type}  "
            f"({self.city_size_var.get()} / {self.wealth_var.get()})"
        )

    def _reroll(self):
        if not self.current_items:
            messagebox.showinfo("Info", "Generate a shop first.")
            return
        pct       = random.randint(10, 30) / 100
        shop_type = self.current_shop_type.get()
        locked    = [i for i in self.current_items if i.get("locked")]
        unlocked  = [i for i in self.current_items if not i.get("locked")]
        n_reroll  = max(1, int(len(unlocked) * pct))
        keep      = random.sample(unlocked, max(0, len(unlocked) - n_reroll))
        weights   = self._get_rarity_weights()
        new_items = generate_shop_items(
            shop_type, len(self.current_items), weights, locked + keep,
            tag_filters=self.active_tag_filters   if self.active_tag_filters   else None,
            tag_excludes=self.excluded_tag_filters if self.excluded_tag_filters else None,
            city_size=self.city_size_var.get(),
            wealth=self.wealth_var.get(),
            mundane_only=self.mundane_only_var.get())
        self.current_items = new_items
        self._populate_table(self.current_items)
        self.status_var.set(
            f"↻  Rerolled ~{int(pct*100)}% of unlocked items ({n_reroll} swapped)"
        )

    def _clear(self):
        if messagebox.askyesno("Clear Shop", "Clear all items?"):
            self.current_items = []
            self._populate_table([])
            self._clear_inspect()
            self.status_var.set("Shop cleared.")

    def _random_name(self):
        self.shop_name_var.set(generate_shop_name(self.current_shop_type.get()))

    def _on_shop_type_change(self, _=None):
        if self._app_settings.get("auto_name_on_change", True):
            self._random_name()

    def _on_select(self, _=None):
        sel = self.tree.selection()
        if not sel:
            return
        iid  = sel[0]
        item = next((i for i in self.current_items if i["name"] == iid), None)
        if item:
            self.selected_row = item
            self._show_inspect(item)

    def _on_double_click(self, _=None):
        sel = self.tree.selection()
        if not sel:
            return
        iid = sel[0]
        for item in self.current_items:
            if item["name"] == iid:
                item["locked"] = not item.get("locked", False)
                break
        self._populate_table(self.current_items)

    # ── Save / Load ───────────────────────────────────────────────────────────
    def _save_shop(self):
        campaign  = self.save_campaign_var.get().strip()
        town      = self.save_town_var.get().strip()
        shop_name = self.shop_name_var.get().strip() or f"{self.current_shop_type.get()} Shop"
        notes     = self.shop_notes_widget.get("1.0", "end").strip() if self.shop_notes_widget else ""

        if not campaign:            messagebox.showerror("Error", "Enter a campaign name."); return
        if not town:
            messagebox.showerror("Error", "Enter a town/location name."); return
        if not self.current_items:
            messagebox.showerror("Error", "Generate a shop first."); return

        con = sqlite3.connect(DB_PATH)
        con.execute("PRAGMA foreign_keys = ON")
        try:
            cur = con.cursor()
            cur.execute("INSERT OR IGNORE INTO campaigns (name) VALUES (?)", (campaign,))
            cur.execute("SELECT id FROM campaigns WHERE name=?", (campaign,))
            camp_id = cur.fetchone()[0]

            # Reuse existing town with the same name in this campaign
            existing_town = cur.execute(
                "SELECT id FROM towns WHERE campaign_id=? AND name=?",
                (camp_id, town)).fetchone()
            if existing_town:
                town_id = existing_town[0]
                cur.execute("UPDATE towns SET city_size=? WHERE id=?",
                            (self.city_size_var.get(), town_id))
            else:
                cur.execute("INSERT INTO towns (campaign_id, name, city_size) VALUES (?,?,?)",
                            (camp_id, town, self.city_size_var.get()))
                town_id = cur.lastrowid

            cur.execute(
                "INSERT INTO shops (town_id, name, shop_type, wealth, last_restocked, notes, "
                "shopkeeper_name, shopkeeper_race, shopkeeper_personality, shopkeeper_appearance) "
                "VALUES (?,?,?,?,?,?,?,?,?,?)",
                (town_id, shop_name, self.current_shop_type.get(),
                 self.wealth_var.get(), datetime.now().isoformat(), notes,
                 self.shopkeeper_name_var.get().strip(),
                 self.shopkeeper_race_var.get().strip(),
                 self.shopkeeper_personality_var.get().strip(),
                 self.shopkeeper_appearance_var.get().strip()))
            shop_id = cur.lastrowid

            for item in self.current_items:
                cur.execute("""INSERT INTO shop_items
                    (shop_id,item_id,name,rarity,item_type,source,page,
                     cost_given,quantity,locked,
                     attunement,damage,properties,mastery,weight,tags,description,table_data,
                     sane_cost,market_price)
                    VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)""",
                    (shop_id, item.get("item_id",""), item["name"],
                     item.get("rarity",""), item.get("item_type",""),
                     item.get("source",""), item.get("page",""),
                     item.get("cost_given",""), item.get("quantity","1"),
                     int(item.get("locked",False)),
                     item.get("attunement",""), item.get("damage",""),
                     item.get("properties",""), item.get("mastery",""),
                     item.get("weight",""), item.get("tags",""),
                     item.get("description",""), item.get("table_data",""),
                     item.get("sane_cost",""), item.get("market_price","")))
            con.commit()
            self.save_status_var.set(f"✔ Saved '{shop_name}' → {campaign} / {town}")
            self._refresh_campaign_list()
        except Exception as e:
            con.rollback()
            messagebox.showerror("Save Failed", f"Could not save shop:\n{e}")
            self.save_status_var.set("⚠ Save failed — no changes written.")
        finally:
            con.close()

    def _refresh_campaign_list(self):
        self.save_tree.delete(*self.save_tree.get_children())
        con = sqlite3.connect(DB_PATH)
        cur = con.cursor()
        for (cid, cname) in cur.execute(
                "SELECT id, name FROM campaigns ORDER BY name"):
            cn = self.save_tree.insert("", "end", iid=f"c{cid}",
                                       text=f"{cname}", values=("",))
            for (tid, tname, tsize) in cur.execute(
                    "SELECT id, name, city_size FROM towns "
                    "WHERE campaign_id=? ORDER BY name", (cid,)):
                tn = self.save_tree.insert(cn, "end", iid=f"t{tid}",
                                           text=f"{tname}",
                                           values=(tsize or "",))
                for (sid, sname, stype, swealth) in cur.execute(
                        "SELECT id, name, shop_type, wealth FROM shops "
                        "WHERE town_id=? ORDER BY name", (tid,)):
                    self.save_tree.insert(tn, "end", iid=f"s{sid}",
                                          text=f"{sname}",
                                          values=(f"{stype} / {swealth}",))
        con.close()

    def _load_selected_shop(self):
        sel = self.save_tree.selection()
        if not sel or not sel[0].startswith("s"):
            messagebox.showinfo("Info", "Select a shop to load.")
            return
        shop_id = int(sel[0][1:])
        con = sqlite3.connect(DB_PATH)
        cur = con.cursor()
        row = cur.execute(
            "SELECT name, shop_type, wealth, notes, "
            "COALESCE(shopkeeper_name,''), COALESCE(shopkeeper_race,''), "
            "COALESCE(shopkeeper_personality,''), COALESCE(shopkeeper_appearance,'') "
            "FROM shops WHERE id=?",
            (shop_id,)).fetchone()
        if not row:
            con.close(); return
        shop_name, shop_type, wealth = row[0], row[1], row[2]
        notes = row[3] or ""
        sk_name, sk_race, sk_personality, sk_appearance = row[4], row[5], row[6], row[7]

        # Resolve campaign and town names so the Save form is pre-filled
        town_row = cur.execute(
            "SELECT t.name, t.city_size, c.name FROM towns t "
            "JOIN campaigns c ON c.id = t.campaign_id "
            "WHERE t.id = (SELECT town_id FROM shops WHERE id=?)",
            (shop_id,)).fetchone()
        town_name, city_size, camp_name = town_row if town_row else ("", "", "")

        items_raw = cur.execute("""
            SELECT item_id,name,rarity,item_type,source,page,
                   cost_given,quantity,locked,
                   attunement,damage,properties,mastery,weight,tags,description,
                   COALESCE(table_data,''), COALESCE(sane_cost,''), COALESCE(market_price,'')
            FROM shop_items WHERE shop_id=?""", (shop_id,)).fetchall()
        con.close()

        self.current_items = [{
            "item_id": r[0], "name": r[1], "rarity": r[2],
            "item_type": r[3], "source": r[4], "page": r[5],
            "cost_given": r[6], "quantity": r[7], "locked": bool(r[8]),
            "attunement": r[9], "damage": r[10],
            "properties": r[11], "mastery": r[12],
            "weight": r[13], "tags": r[14], "description": r[15],
            "table_data": r[16], "sane_cost": r[17], "market_price": r[18],
        } for r in items_raw]

        self.shop_name_var.set(shop_name)
        self.current_shop_type.set(shop_type)
        self.wealth_var.set(wealth)
        if city_size:
            self.city_size_var.set(city_size)
        # Pre-fill save form fields so re-saving is seamless
        self.save_campaign_var.set(camp_name)
        self.save_town_var.set(town_name)
        # Restore notes into the text widget
        if self.shop_notes_widget:
            self.shop_notes_widget.delete("1.0", "end")
            self.shop_notes_widget.insert("1.0", notes)
        # Restore shopkeeper fields
        self.shopkeeper_name_var.set(sk_name)
        self.shopkeeper_race_var.set(sk_race)
        self.shopkeeper_personality_var.set(sk_personality)
        self.shopkeeper_appearance_var.set(sk_appearance)
        self._on_wealth_change()
        self._populate_table(self.current_items)
        self.status_var.set(
            f"Loaded '{shop_name}' ({len(self.current_items)} items)"
        )

    def _delete_selected(self):
        sel = self.save_tree.selection()
        if not sel:
            messagebox.showinfo("Info", "Select an item to delete.")
            return
        if not messagebox.askyesno("Delete", "Delete selected? This cannot be undone."):
            return
        iid = sel[0]
        con = sqlite3.connect(DB_PATH)
        cur = con.cursor()
        if iid.startswith("c"):
            cur.execute("DELETE FROM campaigns WHERE id=?", (int(iid[1:]),))
        elif iid.startswith("t"):
            cur.execute("DELETE FROM towns WHERE id=?", (int(iid[1:]),))
        elif iid.startswith("s"):
            cur.execute("DELETE FROM shops WHERE id=?", (int(iid[1:]),))
        con.commit()
        con.close()
        self._refresh_campaign_list()

    def _export_json(self):
        sel = self.save_tree.selection()
        if not sel or not sel[0].startswith("s"):
            messagebox.showinfo("Info", "Select a shop to export.")
            return
        shop_id = int(sel[0][1:])
        con = sqlite3.connect(DB_PATH)
        cur = con.cursor()
        shop  = cur.execute("SELECT * FROM shops WHERE id=?",
                            (shop_id,)).fetchone()
        items = cur.execute("SELECT * FROM shop_items WHERE shop_id=?",
                            (shop_id,)).fetchall()
        con.close()

        data = {
            "shop": dict(zip(
                ["id","town_id","name","shop_type","wealth",
                 "last_restocked","created_at","notes",
                 "shopkeeper_name","shopkeeper_race",
                 "shopkeeper_personality","shopkeeper_appearance"], shop)),
            "items": [dict(zip(
                ["id","shop_id","item_id","name","rarity","item_type",
                 "source","page","cost_given","quantity",
                 "locked","attunement","damage",
                 "properties","mastery","weight","tags","description"], i))
                for i in items],
        }
        path = filedialog.asksaveasfilename(
            defaultextension=".json",
            filetypes=[("JSON files", "*.json")],
            initialfile=re.sub(r'[\\/:*?"<>|]', "_",
                               data["shop"]["name"]).replace(" ", "_") + ".json")
        if path:
            with open(path, "w") as f:
                json.dump(data, f, indent=2)
            messagebox.showinfo("Exported", f"Shop saved to:\n{path}")

    def _import_json(self):
        path = filedialog.askopenfilename(filetypes=[("JSON files", "*.json")])
        if not path:
            return
        try:
            with open(path) as f:
                data = json.load(f)
        except (json.JSONDecodeError, OSError) as e:
            messagebox.showerror("Import Error", f"Could not read JSON file:\n{e}")
            return
        self.current_items = []
        for i in data.get("items", []):
            self.current_items.append({
                "item_id":     i.get("item_id",""),
                "name":        i.get("name",""),
                "rarity":      i.get("rarity",""),
                "item_type":   i.get("item_type",""),
                "source":      i.get("source",""),
                "page":        i.get("page",""),
                "cost_given":  i.get("cost_given",""),
                "quantity":    i.get("quantity","1"),
                "locked":      bool(i.get("locked",False)),
                "attunement":  i.get("attunement",""),
                "damage":      i.get("damage",""),
                "properties":  i.get("properties",""),
                "mastery":     i.get("mastery",""),
                "weight":      i.get("weight",""),
                "tags":        i.get("tags",""),
                "description": i.get("description",""),
                "table_data":  i.get("table_data",""),
                "sane_cost":   i.get("sane_cost",""),
                "market_price":i.get("market_price",""),
            })
        sdata = data.get("shop", {})
        self.shop_name_var.set(sdata.get("name", "Imported Shop"))
        self.current_shop_type.set(sdata.get("shop_type", "Magic"))
        # Restore wealth level and sync sliders — was missing, leaving UI out of sync
        wealth = sdata.get("wealth", "Average")
        if wealth in WEALTH_DEFAULTS:
            self.wealth_var.set(wealth)
            self._on_wealth_change()
        # Restore notes
        if self.shop_notes_widget:
            self.shop_notes_widget.delete("1.0", "end")
            self.shop_notes_widget.insert("1.0", sdata.get("notes", ""))
        # Restore shopkeeper
        self.shopkeeper_name_var.set(sdata.get("shopkeeper_name", ""))
        self.shopkeeper_race_var.set(sdata.get("shopkeeper_race", ""))
        self.shopkeeper_personality_var.set(sdata.get("shopkeeper_personality", ""))
        self.shopkeeper_appearance_var.set(sdata.get("shopkeeper_appearance", ""))
        self._populate_table(self.current_items)
        self.status_var.set(
            f"Imported {len(self.current_items)} items from JSON."
        )


    # ══════════════════════════════════════════════════════════════════════════
    #  Item Gallery Tab
    # ══════════════════════════════════════════════════════════════════════════

    def _build_gallery_tab(self):
        c = self.colors
        f = self.tab_gallery

        # Left pane
        left = ttk.Frame(f)
        left.pack(side="left", fill="both", expand=True)
        self._gallery_left = left

        # Search / filter bar
        bar = tk.Frame(left, bg=c["hdr"], pady=6)
        bar.pack(fill="x")

        tk.Label(bar, text="◉  Item Gallery",
                 font=("Georgia", 13, "bold"),
                 bg=c["hdr"], fg=c["accent"]).pack(side="left", padx=(10, 16))

        tk.Label(bar, text="⌕", bg=c["hdr"], fg=c["fg"]).pack(side="left")
        self.gallery_search_var = tk.StringVar()
        self.gallery_search_var.trace_add("write", lambda *_: self._gallery_refresh())
        tk.Entry(bar, textvariable=self.gallery_search_var, width=30,
                 bg=c["sel"], fg=c["fg"], insertbackground=c["fg"],
                 relief="flat", font=("Consolas", 9)).pack(side="left", padx=(4, 12))

        tk.Label(bar, text="Rarity:", bg=c["hdr"], fg=c["fg"],
                 font=("Georgia", 9)).pack(side="left")
        self.gallery_rarity_var = tk.StringVar(value="All")
        rarity_opts = ["All", "Mundane", "Common", "Uncommon", "Rare",
                       "Very Rare", "Legendary", "Artifact"]
        ttk.Combobox(bar, textvariable=self.gallery_rarity_var,
                     values=rarity_opts, width=12,
                     state="readonly").pack(side="left", padx=(4, 12))
        self.gallery_rarity_var.trace_add("write", lambda *_: self._gallery_refresh())

        tk.Label(bar, text="Source:", bg=c["hdr"], fg=c["fg"],
                 font=("Georgia", 9)).pack(side="left")
        self.gallery_source_var = tk.StringVar(value="(All)")
        source_combo = ttk.Combobox(bar, textvariable=self.gallery_source_var,
                                     values=_SOURCE_OPTS, width=42,
                                     state="readonly")
        source_combo.pack(side="left", padx=(4, 12))
        self.gallery_source_var.trace_add("write", lambda *_: self._gallery_refresh())

        self.gallery_tag_filters:  set[str] = set()
        self.gallery_tag_excludes: set[str] = set()
        self._gallery_tag_state_vars: dict[str, tk.IntVar] = {}

        tag_outer = tk.Frame(left, bg=c["bg"])
        tag_outer.pack(fill="x", padx=6, pady=(4, 0))

        tag_hdr = tk.Frame(tag_outer, bg=c["sel"], cursor="hand2")
        tag_hdr.pack(fill="x")

        self._gallery_tags_collapsed = tk.BooleanVar(value=True)
        self._gallery_tags_arrow = tk.Label(tag_hdr, text="▶", font=("Consolas", 8),
                                             bg=c["sel"], fg=c["accent"], width=2)
        self._gallery_tags_arrow.pack(side="left", padx=(6, 2))

        tk.Label(tag_hdr, text="Tag Filters",
                 font=("Georgia", 9, "bold"),
                 bg=c["sel"], fg=c["fg"]).pack(side="left", pady=3)
        tk.Label(tag_hdr,
                 text="✓ include  ✗ exclude  (click to cycle)",
                 bg=c["sel"], fg=c["fg"],
                 font=("Georgia", 7, "italic")).pack(side="left", padx=(8, 0))

        ttk.Button(tag_hdr, text="✕ Clear",
                   command=self._gallery_clear_tags).pack(side="right", padx=(0, 4))
        ttk.Button(tag_hdr, text="☑ All",
                   command=self._gallery_select_all_tags).pack(side="right", padx=(0, 2))
        self.gallery_tag_active_lbl = tk.Label(tag_hdr, text="",
                                                bg=c["sel"], fg="#ff9900",
                                                font=("Consolas", 8, "bold"))
        self.gallery_tag_active_lbl.pack(side="right", padx=6)

        self._gallery_tag_body = tk.Frame(tag_outer, bg=c["bg"])

        gtag_canvas_frame = tk.Frame(self._gallery_tag_body, bg=c["bg"], height=120)
        gtag_canvas_frame.pack(fill="x")
        gtag_canvas_frame.pack_propagate(False)

        gtag_canvas = tk.Canvas(gtag_canvas_frame, bg=c["bg"], highlightthickness=0)
        gtag_vsb    = ttk.Scrollbar(gtag_canvas_frame, orient="vertical",
                                    command=gtag_canvas.yview)
        gtag_canvas.configure(yscrollcommand=gtag_vsb.set)
        gtag_vsb.pack(side="right", fill="y")
        gtag_canvas.pack(side="left", fill="both", expand=True)

        gtag_inner = tk.Frame(gtag_canvas, bg=c["bg"])
        gtag_win   = gtag_canvas.create_window((0, 0), window=gtag_inner, anchor="nw")

        def _gtag_inner_configure(e):
            gtag_canvas.configure(scrollregion=gtag_canvas.bbox("all"))
        def _gtag_canvas_configure(e):
            gtag_canvas.itemconfig(gtag_win, width=e.width)
        gtag_inner.bind("<Configure>", _gtag_inner_configure)
        gtag_canvas.bind("<Configure>", _gtag_canvas_configure)

        GTAG_RARITY_COLORS = RARITY_COLORS_MAP
        for cat_name, tags in TAG_CATEGORIES.items():
            self._build_gallery_tag_section(gtag_inner, cat_name, tags, c, GTAG_RARITY_COLORS)

        def _toggle_tag_panel(_=None):
            if self._gallery_tags_collapsed.get():
                self._gallery_tag_body.pack(fill="x")
                self._gallery_tags_arrow.configure(text="▼")
                self._gallery_tags_collapsed.set(False)
            else:
                self._gallery_tag_body.pack_forget()
                self._gallery_tags_arrow.configure(text="▶")
                self._gallery_tags_collapsed.set(True)

        tag_hdr.bind("<Button-1>", _toggle_tag_panel)
        for child in tag_hdr.winfo_children():
            child.bind("<Button-1>", _toggle_tag_panel)

        # ── Pagination bar ──────────────────────────────────────────────────────
        page_bar = tk.Frame(left, bg=c["hdr"])
        page_bar.pack(fill="x", padx=6, pady=(4, 0))

        tk.Label(page_bar, text="Per page:", bg=c["hdr"], fg=c["fg"],
                 font=("Georgia", 9)).pack(side="left", padx=(0, 4))
        per_page_opts = [100, 250, 500, 1000]
        per_page_combo = ttk.Combobox(page_bar,
                                       textvariable=self._gallery_per_page,
                                       values=per_page_opts, width=6,
                                       state="readonly")
        per_page_combo.pack(side="left", padx=(0, 12))
        per_page_combo.bind("<<ComboboxSelected>>",
                             lambda _: self._gallery_go_page(0))

        ttk.Button(page_bar, text="◀ Prev",
                   command=lambda: self._gallery_go_page(self._gallery_page - 1)
                   ).pack(side="left", padx=(0, 4))

        self._gallery_page_lbl = tk.Label(page_bar, text="Page 1 / 1",
                                           bg=c["hdr"], fg=c["accent"],
                                           font=("Georgia", 9, "bold"), width=12)
        self._gallery_page_lbl.pack(side="left", padx=4)

        ttk.Button(page_bar, text="Next ▶",
                   command=lambda: self._gallery_go_page(self._gallery_page + 1)
                   ).pack(side="left", padx=(0, 12))

        self.gallery_count_var = tk.StringVar(value="")
        tk.Label(page_bar, textvariable=self.gallery_count_var,
                 bg=c["hdr"], fg=c["accent"],
                 font=("Georgia", 9, "italic")).pack(side="left")

        tree_frame = ttk.Frame(left)
        tree_frame.pack(fill="both", expand=True, padx=6, pady=4)

        GCOLS   = ("name", "rarity", "type", "source", "value")
        GHDRS   = ("Name", "Rarity", "Type", "Source", "Value")
        GWIDTHS = (290, 95, 200, 80, 100)

        self.gallery_tree = ttk.Treeview(tree_frame, columns=GCOLS,
                                          show="headings", selectmode="browse")
        for col, hdr, w in zip(GCOLS, GHDRS, GWIDTHS):
            self.gallery_tree.heading(col, text=hdr,
                                      command=lambda c=col: self._gallery_sort(c))
            self.gallery_tree.column(col, width=w,
                                     anchor="w" if col in ("name","type") else "center")

        self.gallery_tree.tag_configure("odd",  background=ROW_ODD)
        self.gallery_tree.tag_configure("even", background=ROW_EVEN)
        for rarity, color in GTAG_RARITY_COLORS.items():
            self.gallery_tree.tag_configure(
                rarity.replace(" ", "_"), foreground=color)

        gvsb = ttk.Scrollbar(tree_frame, orient="vertical",
                              command=self.gallery_tree.yview)
        self.gallery_tree.configure(yscrollcommand=gvsb.set)
        self.gallery_tree.pack(side="left", fill="both", expand=True)
        gvsb.pack(side="right", fill="y")
        self.gallery_tree.bind("<<TreeviewSelect>>", self._gallery_on_select)

        self._gallery_inspect_expanded       = False
        self._gallery_inspect_width_collapsed = 400
        self._gallery_inspect_width_expanded  = None

        self.gallery_inspect_panel = tk.Frame(f, bg=c["hdr"])
        self.gallery_inspect_panel.place(relx=1.0, rely=0.0, anchor="ne",
                                          width=self._gallery_inspect_width_collapsed,
                                          relheight=1.0)
        left.pack_configure(padx=(0, self._gallery_inspect_width_collapsed + 6))

        ginsp_hdr = tk.Frame(self.gallery_inspect_panel, bg=c["hdr"])
        ginsp_hdr.pack(fill="x", padx=8, pady=(10, 0))

        tk.Label(ginsp_hdr, text="◈  Item Inspector",
                 font=("Georgia", 11, "bold"),
                 bg=c["hdr"], fg=c["accent"]).pack(side="left")

        self.gallery_expand_btn = tk.Label(
            ginsp_hdr, text="⤢", font=("Georgia", 13),
            bg=c["hdr"], fg=c["fg"], cursor="hand2", padx=4)
        self.gallery_expand_btn.pack(side="right")
        self.gallery_expand_btn.bind("<Button-1>",
                                     lambda e: self._toggle_gallery_inspect_expand())
        self.gallery_expand_btn.bind("<Enter>",
            lambda e: self.gallery_expand_btn.configure(fg=c["accent"]))
        self.gallery_expand_btn.bind("<Leave>",
            lambda e: self.gallery_expand_btn.configure(fg=c["fg"]))

        ttk.Separator(self.gallery_inspect_panel, orient="horizontal").pack(
            fill="x", padx=8, pady=(4, 0))

        ginsp_canvas = tk.Canvas(self.gallery_inspect_panel, bg=c["hdr"],
                                  highlightthickness=0)
        ginsp_vsb    = ttk.Scrollbar(self.gallery_inspect_panel, orient="vertical",
                                     command=ginsp_canvas.yview)
        ginsp_canvas.configure(yscrollcommand=ginsp_vsb.set)
        ginsp_vsb.pack(side="right", fill="y")
        ginsp_canvas.pack(side="left", fill="both", expand=True)

        self.gallery_inspect_frame = tk.Frame(ginsp_canvas, bg=c["hdr"])
        self._ginsp_win = ginsp_canvas.create_window(
            (0, 0), window=self.gallery_inspect_frame, anchor="nw")

        def _ginsp_configure(e):
            ginsp_canvas.configure(scrollregion=ginsp_canvas.bbox("all"))
        def _ginsp_canvas_configure(e):
            ginsp_canvas.itemconfig(self._ginsp_win, width=e.width)
        self.gallery_inspect_frame.bind("<Configure>", _ginsp_configure)
        ginsp_canvas.bind("<Configure>", _ginsp_canvas_configure)

        tk.Label(self.gallery_inspect_frame,
                 text="Search and select an item to inspect.",
                 bg=c["hdr"], fg=c["fg"],
                 font=("Georgia", 9, "italic")).pack(pady=20, padx=10)

        self._gallery_refresh()

    # ── Gallery expand / collapse ─────────────────────────────────────────────
    def _toggle_gallery_inspect_expand(self):
        expanding = not self._gallery_inspect_expanded
        if self._gallery_inspect_width_expanded is None or expanding:
            self.update_idletasks()
            win_w = self.winfo_width()
            self._gallery_inspect_width_expanded = max(600, int(win_w * 0.56))
        w = (self._gallery_inspect_width_expanded if expanding
             else self._gallery_inspect_width_collapsed)
        self.gallery_inspect_panel.place_configure(width=w)
        self._gallery_left.pack_configure(padx=(0, w + 6))
        self._gallery_inspect_expanded = expanding
        self.gallery_expand_btn.configure(text="⤡" if expanding else "⤢")

    # ── Gallery tag sections (3-state: neutral / include / exclude) ─────────────
    def _build_gallery_tag_section(self, parent, cat_name: str, tags: list[str],
                                    c: dict, rarity_colors: dict):
        section = tk.Frame(parent, bg=c["bg"],
                           highlightbackground=c["sel"], highlightthickness=1)
        section.pack(fill="x", padx=4, pady=3)
        collapsed = tk.BooleanVar(value=True)

        hdr = tk.Frame(section, bg=c["sel"], cursor="hand2")
        hdr.pack(fill="x")
        arrow_lbl = tk.Label(hdr, text="▶", font=("Consolas", 8),
                              bg=c["sel"], fg=c["accent"], width=2)
        arrow_lbl.pack(side="left", padx=(6, 2))
        tk.Label(hdr, text=cat_name, font=("Georgia", 9, "bold"),
                 bg=c["sel"], fg=c["fg"]).pack(side="left", pady=4)
        count_var = tk.StringVar(value="")
        tk.Label(hdr, textvariable=count_var, bg=c["sel"], fg="#ff9900",
                 font=("Consolas", 8)).pack(side="right", padx=8)
        body = tk.Frame(section, bg=c["bg"])

        STATE_FG  = {0: c["fg"],    1: "#1eff00", 2: "#ff4444"}
        STATE_BG  = {0: c["bg"],    1: "#0d1f0d", 2: "#1f0d0d"}
        STATE_PFX = {0: "  ",       1: "✓ ",      2: "✗ "}

        def _refresh_count():
            n_inc = sum(1 for t in tags
                        if self._gallery_tag_state_vars.get(t, tk.IntVar()).get() == 1)
            n_exc = sum(1 for t in tags
                        if self._gallery_tag_state_vars.get(t, tk.IntVar()).get() == 2)
            parts = []
            if n_inc: parts.append(f"{n_inc} incl")
            if n_exc: parts.append(f"{n_exc} excl")
            count_var.set(" / ".join(parts))

        def _toggle_section(_=None):
            if collapsed.get():
                body.pack(fill="x", padx=8, pady=(4, 6))
                arrow_lbl.configure(text="▼")
                collapsed.set(False)
            else:
                body.pack_forget()
                arrow_lbl.configure(text="▶")
                collapsed.set(True)

        hdr.bind("<Button-1>", _toggle_section)
        for child in hdr.winfo_children():
            child.bind("<Button-1>", _toggle_section)

        cols = 4
        for idx, tag in enumerate(tags):
            var = tk.IntVar(value=0)
            self._gallery_tag_state_vars[tag] = var
            btn_ref: list = []

            def _cycle(t=tag, v=var, br=btn_ref, rf=_refresh_count):
                new_state = (v.get() + 1) % 3
                v.set(new_state)
                self.gallery_tag_filters.discard(t)
                self.gallery_tag_excludes.discard(t)
                if new_state == 1:
                    self.gallery_tag_filters.add(t)
                elif new_state == 2:
                    self.gallery_tag_excludes.add(t)
                if br:
                    br[0].configure(
                        text=STATE_PFX[new_state] + t,
                        fg=STATE_FG[new_state],
                        bg=STATE_BG[new_state],
                    )
                rf()
                self._update_gallery_tag_summary()
                self._gallery_refresh()

            btn = tk.Button(
                body,
                text=STATE_PFX[0] + tag,
                command=_cycle,
                fg=STATE_FG[0], bg=STATE_BG[0],
                activeforeground=c["accent"],
                activebackground=c["sel"],
                relief="flat", bd=0,
                font=("Georgia", 8),
                anchor="w", padx=2,
            )
            btn_ref.append(btn)
            btn.grid(row=idx // cols, column=idx % cols, sticky="w", padx=2, pady=1)

    def _update_gallery_tag_summary(self):
        if not hasattr(self, "gallery_tag_active_lbl"):
            return
        n_inc = len(self.gallery_tag_filters)
        n_exc = len(self.gallery_tag_excludes)
        parts = []
        if n_inc: parts.append(f"{n_inc} included")
        if n_exc: parts.append(f"{n_exc} excluded")
        self.gallery_tag_active_lbl.configure(text=" / ".join(parts))

    def _gallery_tag_toggle(self, tag: str, var: tk.BooleanVar, refresh_count_fn=None):
        pass

    def _gallery_clear_tags(self):
        self.gallery_tag_filters.clear()
        self.gallery_tag_excludes.clear()
        for var in self._gallery_tag_state_vars.values():
            var.set(0)
        self._repaint_gallery_tag_buttons()
        self.gallery_tag_active_lbl.configure(text="")
        self._gallery_refresh()

    def _gallery_select_all_tags(self):
        self.gallery_tag_filters.clear()
        self.gallery_tag_excludes.clear()
        for tag, var in self._gallery_tag_state_vars.items():
            var.set(1)
            self.gallery_tag_filters.add(tag)
        self._repaint_gallery_tag_buttons()
        self._update_gallery_tag_summary()
        self._gallery_refresh()

    def _repaint_gallery_tag_buttons(self):
        """Repaint all gallery tag buttons to match their current IntVar state."""
        STATE_FG  = {0: self.colors["fg"], 1: "#1eff00", 2: "#ff4444"}
        STATE_BG  = {0: self.colors["bg"], 1: "#0d1f0d", 2: "#1f0d0d"}
        STATE_PFX = {0: "  ",             1: "✓ ",      2: "✗ "}
        for tag, var in self._gallery_tag_state_vars.items():
            s = var.get()
            for widget in self._iter_gallery_tab_buttons():
                txt = widget.cget("text")
                if len(txt) >= 2 and txt[2:] == tag:
                    widget.configure(
                        text=STATE_PFX[s] + tag,
                        fg=STATE_FG[s],
                        bg=STATE_BG[s],
                    )
                    break

    def _iter_gallery_tab_buttons(self):
        """Yield tk.Button widgets inside the gallery tag canvas."""
        def _recurse(w):
            if isinstance(w, tk.Button):
                yield w
            for child in w.winfo_children():
                yield from _recurse(child)
        if hasattr(self, "tab_gallery"):
            yield from _recurse(self.tab_gallery)

    def _gallery_go_page(self, page: int):
        """Navigate to a specific page (clamped to valid range) and re-render."""
        total = len(self._gallery_all_results)
        per   = max(1, int(self._gallery_per_page.get()))
        max_page = max(0, math.ceil(total / per) - 1)
        self._gallery_page = max(0, min(page, max_page))
        self._gallery_render_page()

    def _gallery_render_page(self):
        """Render the current page slice into the treeview and update the page label."""
        total   = len(self._gallery_all_results)
        per     = max(1, int(self._gallery_per_page.get()))
        page    = self._gallery_page
        max_page = max(0, math.ceil(total / per) - 1)
        start   = page * per
        end     = start + per

        page_slice = self._gallery_all_results[start:end]
        self._gallery_results = page_slice

        total_pages = max(1, math.ceil(total / per))
        self._gallery_page_lbl.configure(
            text=f"Page {page + 1} / {total_pages}")
        self.gallery_count_var.set(
            f"{total:,} items  •  {start + 1}–{min(end, total)} shown")

        self.gallery_tree.delete(*self.gallery_tree.get_children())
        for idx, item in enumerate(page_slice):
            rnorm  = normalize_rarity(item.get("Rarity", ""))
            r_tag  = rnorm.replace(" ", "_")
            parity = "odd" if idx % 2 == 0 else "even"
            self.gallery_tree.insert("", "end",
                iid=f"g_{idx}",
                values=(
                    item.get("Name", ""),
                    item.get("Rarity", "—"),
                    item.get("Type", "—"),
                    item.get("Source", "—"),
                    item.get("Value", "—") or "—",
                ),
                tags=(parity, r_tag),
            )

    def _gallery_refresh(self):
        q       = self.gallery_search_var.get().strip().lower()
        rfilter = self.gallery_rarity_var.get()
        raw_source = self.gallery_source_var.get()
        if raw_source and raw_source != "(All)":
            sfilter = raw_source.split(" — ")[0].strip().lower()
        else:
            sfilter = ""

        results = []
        for item in ALL_ITEMS_FLAT:
            if q and q not in item.get("Name", "").lower():
                continue
            if rfilter != "All":
                if normalize_rarity(item.get("Rarity", "")) != rfilter.lower():
                    continue
            if sfilter and item.get("Source", "").lower() != sfilter:
                continue
            if self.gallery_tag_excludes or self.gallery_tag_filters:
                item_tags = {t.strip() for t in item.get("Tags", "").split(",") if t.strip()}
                if self.gallery_tag_excludes and (item_tags & self.gallery_tag_excludes):
                    continue
                if self.gallery_tag_filters and not (item_tags & self.gallery_tag_filters):
                    continue
            results.append(item)

        # Sort
        col = self._gallery_sort_col
        rev = not self._gallery_sort_asc
        if col == "name":
            results.sort(key=lambda x: x.get("Name", "").lower(), reverse=rev)
        elif col == "rarity":
            results.sort(key=lambda x: (rarity_rank(x.get("Rarity", "")),
                                         x.get("Name", "").lower()), reverse=rev)
        elif col == "source":
            results.sort(key=lambda x: x.get("Source", "").lower(), reverse=rev)
        elif col == "value":
            results.sort(key=lambda x: parse_given_cost(x.get("Value", "")) or 0, reverse=rev)
        else:
            results.sort(key=lambda x: x.get("Name", "").lower(), reverse=rev)

        self._gallery_all_results = results
        self._gallery_page = 0
        self._gallery_render_page()

    def _gallery_sort(self, col: str):
        if self._gallery_sort_col == col:
            self._gallery_sort_asc = not self._gallery_sort_asc
        else:
            self._gallery_sort_col = col
            self._gallery_sort_asc = True
        self._gallery_refresh()

    def _gallery_on_select(self, _=None):
        sel = self.gallery_tree.selection()
        if not sel:
            return
        iid = sel[0]
        try:
            idx = int(iid.split("_")[1])
            raw = self._gallery_results[idx]
        except (IndexError, ValueError):
            return

        rarity_raw = raw.get("Rarity", "")
        mkt = generate_market_price(rarity_raw)
        item = {
            "name":         raw.get("Name", ""),
            "rarity":       rarity_raw,
            "item_type":    raw.get("Type", ""),
            "source":       raw.get("Source", ""),
            "page":         raw.get("Page", ""),
            "item_id":      raw.get("Item ID", ""),
            "cost_given":   raw.get("Value", ""),
            "attunement":   raw.get("Attunement", ""),
            "damage":       raw.get("Damage", ""),
            "properties":   raw.get("Properties", ""),
            "mastery":      raw.get("Mastery", ""),
            "weight":       raw.get("Weight", ""),
            "tags":         raw.get("Tags", ""),
            "description":  raw.get("Text", ""),
            "table_data":   raw.get("Table", ""),
            "quantity":     "",
            "locked":       False,
            "_gallery":     True,
            "sane_cost":    raw.get("Sane_Cost", ""),
            "market_price": str(mkt) if mkt is not None else "",
        }

        for w in self.gallery_inspect_frame.winfo_children():
            w.destroy()

        _real_frame = self.inspect_frame
        self.inspect_frame = self.gallery_inspect_frame
        self._render_inspect_collapsed(item)
        self.inspect_frame = _real_frame


# ══════════════════════════════════════════════════════════════════════════════
#  Entry point
# ══════════════════════════════════════════════════════════════════════════════
def main():
    app = ShopApp()
    app.mainloop()

if __name__ == "__main__":
    main()
