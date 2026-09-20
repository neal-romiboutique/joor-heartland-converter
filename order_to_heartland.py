#!/usr/bin/env python3
"""
JOOR -> Heartland Retail converter for Romi Boutique.  (v7)

Phase 1 (grids):  python3 joor_to_heartland.py grids JOOR_ORDER.csv|.xlsx [-o grids.xlsx]
    One row per (vendor, style, color) in Heartland's official grid-import
    template schema, Generate Items = true, sizes completed per store rules.

Phase 2 (po):     python3 joor_to_heartland.py po JOOR_ORDER.csv|.xlsx ITEMS_EXPORT.csv [-o po.xlsx]
    Joins Heartland's post-import grid/items export back to the JOOR order
    (grid # + size -> Item #) and emits the PO import file.

Accepts BOTH JOOR export formats, auto-detected:
  A) Basic Export (long): csv/tsv/txt/xlsx, header row 1, snake_case columns,
     one row per style/color/size.
  B) Order-to-Size (wide): formatted xlsx PO, metadata block on top, one row
     per style/color with sizes as columns; may contain embedded product
     images, which the grids phase extracts to <output>_images/.

Store conventions encoded below; edit CONFIG to change.
"""
import argparse, csv, math, os, re, sys, time, unicodedata
import requests
from decimal import Decimal, ROUND_CEILING
from pathlib import Path

# ----------------------------- CONFIG ---------------------------------------
MARGIN = Decimal("2.4")          # retail = MARGIN x unit cost when JOOR has no MSRP
VENDOR_MAP = {                   # JOOR seller_account -> exact Heartland vendor name
    # Matched through _vendor_key, so casing/accents/punctuation don't matter.
    "Mother Denim": "Mother",    # JOOR trades as Mother Denim; Heartland has "Mother"
}
# Known Heartland vendors (Purchasing -> Vendors export), used to resolve the
# exact existing vendor Name instead of guessing one. Re-export and regenerate
# this list whenever new vendors are added in Heartland.
KNOWN_VENDOR_NAMES = [
    '143',
    '27 Miles',
    '3 X 1',
    '360 Cashmere',
    '382',
    '525',
    '7 For All Mankind',
    '9SEED',
    'A Shirt Thing',
    'A.L.C',
    'AGOLDE',
    'AMERICAN COLORS',
    'APPARIS',
    'AS 98',
    'AS BY DF',
    'AS98',
    'ASKK NY',
    'ATM',
    'Abigail Jayne Design',
    'Adina Reyter',
    'Alemais',
    'Alexa Leigh',
    'Alice & Olivia',
    'Alice & Trixie',
    'Alix of Bohemia',
    'All Time Favorite',
    'Aloha',
    'Amalia',
    'Amanda Uprichard',
    'American Vintage',
    'Amo',
    'Ancient Greek Sandals',
    'Angel Wings',
    'Anonym',
    'Ansaldo',
    'Apres Jewelry',
    'Aquarius Cocktail',
    'Aquatalia',
    'Autumn Cashmere',
    'Avenue Montaigne',
    'Aviator Nation',
    'Azul',
    'B. Belt',
    'B.Belt',
    'BA&SH',
    'BALLIER',
    'BB Dakota',
    'BE JE',
    'BLOOM JEWELRY',
    'BOHO',
    'Banjanan',
    'Be Love',
    'Beek',
    'Bell',
    'Bella Dahl',
    'Bella Sleep + Spa',
    'Bevy Flog',
    'Bibi Lou',
    'Blonde No 8',
    'Brazeau Tricot',
    'Britta',
    'Brochu Walker',
    'Brodie',
    'Bronx And Banco',
    'Bubish',
    'CALLAHAN',
    'CALYPSO',
    'CARDEN AVENUE JEWELRY',
    'CATHERINE GEE',
    'CHINTI AND PARKER',
    'CHRISTY LYNN',
    'CINO',
    'CIRCOLO',
    'COLORI',
    'CRUSH',
    'Caballero',
    'Caddis',
    'California Tailor',
    'Calleen Cordero',
    'Cami',
    'Cami NYC',
    'Campomaggi',
    'Cara Cara Clothing',
    'Carrie Forbes',
    'Caterina Lucchi',
    'Central Park WEST',
    'Chan Luu',
    'Charo Ruiz Ibiza',
    'Chaser',
    'Chie Mihara',
    'Christopher Fischer',
    'Ciao Lucia',
    'Cinq A Sept',
    'Citizens',
    'Clare V',
    'Clare Vivier',
    'Claris Virot',
    'Cleobella',
    'Coclico',
    'Colorush',
    'Commando',
    'Corazon Playero',
    'Cordani',
    'Cosabella',
    'Cotton Citizen',
    'Cowboy Belts',
    'Cupcakes & Cashmere',
    'Current Elliot',
    'DANIEL RAHIM',
    'DELACY',
    'DL1961',
    'DOEN',
    'DONNI.',
    'Dear Frances',
    'Denimist',
    'Derek Lam',
    'Design History',
    'Devon Dowd',
    'Diane Von Furstenberg',
    'Dolce Vita',
    'Doma',
    'Dragon Diffusion',
    'Drew',
    'EF Collection',
    'EJ Jewelry',
    'ERIN GRAY',
    'Eberjey',
    'Ecru',
    'Elaine Kim',
    'Electric & Rose',
    'Elizabeth and James',
    'Ella Moss',
    'Embrazio',
    'Enza Costa',
    'Equipment',
    'Estheme Cashmere',
    'Eterne',
    'Eugenia Kim',
    'FARM Rio',
    'FIGUE',
    'FITE',
    'FRANK & EILEEN',
    'FRANK & EILEEN TEE LAB',
    'FREECITY',
    'FUN SOCKS',
    'Faliero Sarti',
    'Feel The Piece',
    'Felicite',
    'Fig & Bella',
    'Florabella',
    'Frame',
    'Freda Salvador',
    'Freya',
    'G1',
    'GANNI',
    'GBTSO',
    'GILNER FARRAR',
    'GOGO',
    'GRLFRND',
    'Galadriel Mattei',
    'Generation Love',
    'Gentle Fawn',
    'Giles and Brothers',
    'Global Girls',
    'Go Silk',
    'Goddis',
    'Gold Hawk',
    'Goldsign',
    'Good hYouman',
    'Groceries Apparel',
    'Guest In Residence',
    'Gypsy',
    'HAPPY SHEEP',
    'HASBEENS',
    'Halston Heritage',
    'Hammitt',
    'Hartford',
    'Haute Hippie',
    'Havaiana',
    'Helene',
    'Hemant & Nandita',
    'Hudson',
    'Hula Hoops',
    "I Stole My Boyfriend's Shirt",
    'ICONS',
    'IKKS',
    'IRO',
    'Inhabit',
    'Isabel Marant',
    'J-ALL',
    'JBrand',
    'JOHN ELLIOTT',
    'Jakett',
    'Jakett New York',
    'James Perse',
    'Jane Hollinger',
    'Jane Tran',
    'Janessa Leone',
    'Jen R',
    'Jennifer Haley',
    'Jennifer Zeuner',
    'Jenny Bird',
    'Jet',
    'Joelle Casady',
    'John & Jen',
    'John & Jenn',
    'Joie',
    'Joy Dravecky',
    'Joy Li',
    'Julie Shenkman',
    'Jumper 1234',
    'June',
    'June 7.2',
    'Junk Food',
    'Jurate Brown',
    'KERO',
    'KERRI ROSENTHAL',
    'KMJ',
    'KNTIWIT',
    'Kane',
    'Kelly Wearstler',
    'Kersh',
    'Kim Vogel',
    'Kim White',
    'Kobo',
    'Krisa',
    'Kujten',
    "L'Agence",
    'LAFCO',
    'LBLC',
    'LE SUPERBE',
    'LNA',
    'LOLA & SOPHIE',
    'LOVE SAM',
    'LOVE TANJANE',
    'La Made',
    'La Natura',
    'La Prestic Ouiston',
    'Lana Bilzerian',
    'Lanston',
    'Lauren & Gracia',
    'Lauren Moshi',
    'Lavender Crush',
    'Le Bon Shoppe',
    'Le Divina',
    'Le Jean',
    'Leallo',
    'Leighelena',
    'Letarte',
    'Level 9',
    'Lilla P',
    'Line',
    'Lingua Franca',
    'Lisa Freede',
    'Lisa Yang',
    'Local Beach',
    'Loeffler Randall',
    'Lola Cruz',
    'Love & Bambii',
    'Love The Label',
    'LoveShackFancy',
    'Loyd/Ford',
    'Lumilla',
    'Luv AJ',
    'Lynn Tallerico',
    'MARIE OLIVER',
    'MARIET MENTAAL',
    'ME & K',
    'MES DEMOISELLES',
    'MILLY CABANA',
    'MOLLY G',
    'Mabe',
    'Madeleine Thompson',
    'Madeworn',
    'Maiami',
    'Maison Hotel',
    'Maison Montagut',
    'Marcia Moran',
    'Marie Chavez',
    'Marit Rae',
    "Mason's",
    'Matta',
    'Mauritius Leather',
    'Maya Bauman',
    'McGuire',
    'Me369',
    'Mercado Global',
    'Mexicana',
    'Michael Lauren',
    'Michael Stars',
    'Mikoh',
    'Milly',
    'Minnie Rose',
    'Mint & Rose',
    'Misa',
    'Miz Mooz',
    'Momoni',
    'Monrow',
    'Mother',
    'Mou',
    'Moussy',
    'Muuba',
    'NICKHO REY',
    'Naadam',
    'Naghedi',
    'Naom',
    'Nation',
    'Native Gem',
    'Never A Wallflower',
    'Nibi MTK',
    'Nili Lotan',
    'Nothing Personal',
    'Nsf',
    'Nubra',
    'ONE GREY DAY',
    'Oats Cashmere',
    'Odette',
    'Olivia Dar',
    'Oncept',
    'Oxbow',
    'PARKER NY',
    'PETE',
    'PHILANTHROPY',
    'PINE CASHMERE',
    'PJ SALVAGE',
    'Paige',
    'Pam & Gela',
    'Paradigm',
    'Parrish',
    'Paul Green',
    'Paula Carvalho',
    'Paula Rosen',
    'Peace Love World',
    'Pharaoh',
    'Pliers & String',
    'Poupette St. Barth',
    'Puntovita',
    'Q',
    'Quilted Koala',
    'RACHEL PALLY',
    'RAQUEL ALLEGRA',
    'ROCK YOUR RADIANCE',
    'ROI',
    'ROSE CARMINE',
    'Rachel Gilbert',
    'Rag & Bone',
    'Rails',
    'Ramona & Ruth',
    'Ramy Brook',
    'Re/Done',
    'Rebecca Scott',
    'Rebels',
    'Retrofete',
    'Richiami',
    'Roam',
    'Robin Haley',
    'Romi',
    'SACCHARINE NY',
    'SALONI',
    'SALTWATER LUXE',
    'SANS FAFF',
    'SANTORE JEWELRY',
    'SARAH DUNN',
    'SEYCHELLES',
    'SOFTWAVES',
    'SOH',
    'SOLID & STRIPED',
    'SPRWMN',
    'STATESIDE',
    'Sablyn',
    'Salt Water LUXE',
    'Sanctuary',
    'Satya Jewelry',
    'Schoen by Yu',
    'Sea',
    'Simon Miller',
    'Skull Cashmere',
    'Soft Joie',
    'Soia & Kyo',
    'Sol Los Angeles',
    'Sold Out NYC',
    'Solei Sea',
    'Sonia & Kyo',
    'Soulbyrd',
    'Splendid',
    'Staud',
    'Streets Ahead',
    'Studio M',
    'Subtle Luxury',
    'Sundry',
    'Susana Monaco',
    'Suzie Kondi',
    'TAI',
    'TKees',
    'TRINA TURK',
    'TWP',
    'TYLHO',
    'Temptation Positano',
    'The Jacksons',
    'Theia',
    'Tilo',
    'Toast',
    'Travata',
    'Tyler Jacobs',
    'Tysa',
    'Ulla Johnson',
    'Unsweetened NY',
    'VICTORIA CUNNINGHAM',
    'Van Palma',
    'Vanzetti',
    'Veronica Beard',
    'Viereck',
    'Vince',
    'Vipul Shah',
    'Vita Fede',
    'Voyage et Cie',
    'W.Cashmere',
    'WESRF',
    'Wandler',
    'White & Warren',
    'Wilt',
    'Wooden Ships',
    'Wyeth',
    'YFB',
    'Yigal Azrouel',
    'Yosi Samra',
    'Yuzefi',
    'Zadig & Voltaire',
    'Zoe Chicco',
    'Zofia Day',
    'ashley ashoff',
    'm&m bling',
    'r & r',
    'rjp',
    'road 22',
    'signorelli',
    'swone',
    'tangerine',
]
SEASON_FROM = "linesheet"        # JOOR column used for Season (long format only)

# Controlled-vocabulary fields on the grid template: rendered as in-cell
# dropdowns so the human reviewer picks from these instead of typing free text.
DROPDOWN_LISTS = {
    "Shopify Category": ["Apparel", "Bags", "Shoes", "Jewelry", "Accessories"],
    "Shopify Color": ["White", "Red", "Burgundy", "Pink", "Orange", "Yellow", "Green",
                      "Blue", "Navy", "Ivory", "Cream", "Tan", "Grey", "Brown", "Black",
                      "Gold", "Multi", "Clear", "Silver", "Yellow Gold", "White Gold",
                      "Two Tone", "Black Diamond", "Emerald", "Turquoise", "Rose Gold",
                      "Pearl", "Labradorite"],
    "Active fit guide": ["1", "2", "3", "4", "5"],
    "Category": ["Dresses", "Denim", "Tops", "T-Shirts", "Sweaters", "Loungewear",
                "Skirts", "Jackets", "Pants", "Shorts", "Handbags", "Shoes", "Scarves",
                "Belts", "Jewelry", "Fine jewelry", "Lingerie", "Sleepwear", "Fragrance",
                "Other Accessories", "Greeting Cards", "Hats & Gloves", "Other"],
}
# Heartland official grid-import template headers (84 cols), verbatim:
TEMPLATE_HEADERS = ['Item Grid #','Description','Long Description','Default Cost','Original Price','Price','Primary Vendor','Color','Category','Size','Style Name','Style Number','Season','Taxable?','Sell on Shopify','Shopify Product Type','Shopify Tags','Shopify Category','Shopify Color','Active fit guide','Tags','Dimension 1','Dimension 1 Values','Dimension 2','Dimension 2 Values','Dimension 3','Dimension 3 Values','Active','Generate Items','Item #','Item Description','Item Default Cost','Item Tags','Item Dynamic Margin','Item Original Price','Item Current Price','Item Active?','Item Track Inventory?','Item Long Description','Item Prompt for description?','Item Prompt for price?','Item Primary Image','Item Use dynamic margin?','Item Primary Vendor','Item Color','Item Category','Item Size','Item Style Name','Item Style Number','Item Season','Item Taxable?','Item Sell on Shopify','Item Shopify Product Type','Item Shopify Tags','Item Shopify Category','Item Shopify Color','Item Active fit guide','Item Weight','Item Weight Unit','Item Width','Item Width Unit','Item Height','Item Height Unit','Item Depth','Item Depth Unit','Item romi boutique Reorder Point','Item romi boutique Target Qty','Item romi boutique Bin Location','Item Shopify Reorder Point','Item Shopify Target Qty','Item Shopify Bin Location','Item Max Retail Reorder Point','Item Max Retail Target Qty','Item Max Retail Bin Location','Item Garmentory Reorder Point','Item Garmentory Target Qty','Item Garmentory Bin Location','Item All Locations Reorder Point','Item All Locations Target Qty','Item All Locations Bin Location','Item Vendor Details Item #','Item Vendor Details Default Cost','Item Vendor Details #','Item Images URL']

# Non-size headers in the wide (Order-to-Size) line table. Anything else with a
# non-empty header between "Country of Origin" and the money columns is a size.
WIDE_NON_SIZE = {"style image", "style name", "style number", "color", "color code",
                 "color comment", "style comment", "materials", "fabrication",
                 "country of origin", "item discount", "units"}
WIDE_NON_SIZE_PREFIXES = ("sugg. retail", "sugg retail", "wholesale", "total")

# ------------------------- text utilities -----------------------------------
def title_caps(s: str) -> str:
    """Word-wise title case, capitalizing after spaces and hyphens,
    keeping letters after apostrophes lowercase (Isabel's, Off-White)."""
    def cap_word(w):
        parts = re.split(r"([-/])", w)   # keep the delimiters, capitalize each chunk
        return "".join(p[:1].upper() + p[1:].lower() if p and p not in "-/" else p
                      for p in parts)
    return " ".join(cap_word(w) for w in s.strip().split())

def fix_mojibake(s: str) -> str:
    """Repair UTF-8 text that was mis-decoded as Mac Roman / cp1252 upstream
    (e.g. 'D√¥en' -> 'Dôen'). Best effort; returns input if repair fails."""
    if not any(ch in s for ch in ("√", "Ã", "â", "€")):
        return s
    for enc in ("mac_roman", "cp1252"):
        try:
            fixed = s.encode(enc).decode("utf-8")
            if sum(ord(c) > 127 for c in fixed) < sum(ord(c) > 127 for c in s):
                return fixed
        except (UnicodeEncodeError, UnicodeDecodeError):
            continue
    return s

_W_SLASH_RE = re.compile(r"(?i)\bW/")

def normalize_color(raw: str) -> str:
    """Title-case a color and join multi-color names with a bare slash
    ('BLUE/BLACK/GREEN' -> 'Blue/Black/Green'). 'W/' is shorthand for 'with',
    not a color divider, so it's protected.

    The slash is the store's standard: of the 17,573 grids in Heartland,
    1,358 carry a '/' in the color and only 40 carry an '&' -- and most of
    those 40 are genuine vendor color NAMES ('Sweet & Sassy', 'Diamonds &
    Rose Gold'), not separators. Between 2026-08-21 and 2026-09-15 this
    function wrote ' & ' instead; grids from that window are matched by
    _color_match_key, not by rewriting them."""
    s = title_caps(fix_mojibake(raw))
    s = _W_SLASH_RE.sub("\x00", s)          # shield W/ from the split
    s = re.sub(r"\s*/\s*", "/", s)
    return s.replace("\x00", "W/")

def _color_match_key(color: str) -> str:
    """Canonical form for COMPARING two colors -- never for writing one.

    Heartland holds the same physical color under every separator style,
    because each was the house convention at some point and some grids were
    typed by hand: 'BLACK/CREAM', 'BLACK / CREAM', 'Black & Cream' and
    'Black/Cream' all coexist in the live catalog (Generation Love H22369 has
    two grids differing only by the spaces around the slash). Enumerating
    variants missed the spaced form; collapsing the separator on BOTH sides of
    the comparison catches all of them at once. A real ampersand inside a
    color name ('Sweet & Sassy') canonicalizes the same way on both sides, so
    it still only ever matches itself."""
    return re.sub(r"\s*[&/]\s*", "/", str(color).strip().upper())
    
def clean_size(raw) -> str:
    """Strip a leading 'SIZE' label some vendors include in JOOR (e.g.
    'SIZE 32' -> '32', 'Size34' -> '34'), regardless of sizing convention.
    Applied wherever a raw size token first enters the pipeline."""
    s = str(raw).strip()
    return re.sub(r"(?i)^size\s*", "", s).strip()

def _vendor_key(s: str) -> str:
    """Normalize a vendor string for matching: strip accents/case/punctuation
    so 'Dôen', 'DOEN', and 'Do\u0308en' all collapse to the same key."""
    s = unicodedata.normalize("NFKD", s)
    s = "".join(c for c in s if not unicodedata.combining(c))
    s = s.upper().replace("&", "AND")
    s = re.sub(r"[.\',/]", " ", s)
    return re.sub(r"[^A-Z0-9]+", " ", s).strip()

# Populated below, once VENDOR_ID_MAP is available, and again from the live
# vendor list at the start of every run (see _rebuild_vendor_lookups).
_VENDOR_LOOKUP = {}

# VENDOR_MAP is keyed through _vendor_key so an override fires whatever casing
# or punctuation JOOR happens to send ("MOTHER DENIM", "Mother Denim.").
_VENDOR_MAP_BY_KEY = {_vendor_key(k): v for k, v in VENDOR_MAP.items()}

def _vendor_near_matches(key):
    """Known Heartland vendors that differ from `key` only by trailing whole
    words: JOOR's "Mother Denim" -> Heartland's "Mother". WHOLE tokens only,
    so 'CAMI' can never match 'CAMILLA'.

    Deliberately returns every candidate instead of picking one. The live list
    contains real, distinct vendors that differ by exactly one trailing word --
    Cami / Cami NYC, Milly / Milly Cabana, Jakett / Jakett New York, Frank And
    Eileen / Frank And Eileen Tee Lab -- so guessing between them would create
    orders against the wrong vendor. The exact-key lookup runs first, so a name
    that IS one of those pairs resolves on its own and never reaches here."""
    toks = key.split()
    if not toks:
        return []
    out = []
    for k, name in _VENDOR_LOOKUP.items():
        kt = k.split()
        if kt != toks and (kt[:len(toks)] == toks or toks[:len(kt)] == kt):
            out.append(name)
    return sorted(out)

def normalize_vendor(raw: str) -> str:
    """Resolve raw JOOR vendor text to the EXACT existing Heartland vendor
    Name (matches how Heartland's import screen looks up vendors) rather than
    guessing a new string, which risks creating a duplicate vendor record."""
    raw = fix_mojibake(raw)
    name = raw.split(" (")[0]                                   # drop "(shopdoen.com)"
    key = _vendor_key(name)
    if key in _VENDOR_MAP_BY_KEY:
        return _VENDOR_MAP_BY_KEY[key]
    hit = _VENDOR_LOOKUP.get(key)
    if hit:
        return hit
    near = _vendor_near_matches(key)
    if len(near) == 1:
        print(f"  NOTE: JOOR vendor '{name}' matched Heartland vendor "
              f"'{near[0]}' (differs only by a trailing word).")
        return near[0]
    guess = unicodedata.normalize("NFKD", name)
    guess = "".join(c for c in guess if not unicodedata.combining(c)).upper().strip()
    if near:
        print(f"  WARNING: vendor '{name}' is ambiguous -- it could be any of: "
              f"{', '.join(near)}. Using '{guess}' as a guess. Type the exact "
              f"Heartland name in the Vendor override box, or add it to "
              f"VENDOR_MAP, before importing.")
    else:
        print(f"  WARNING: vendor '{name}' not found in the known Heartland vendor "
              f"list; using '{guess}' as a guess. Check Purchasing -> Vendors before "
              f"importing, or add an override to VENDOR_MAP, to avoid creating a "
              f"duplicate vendor record.")
    return guess
    
# ------------------------- Heartland API (live dedupe check) ----------------
def _load_heartland_env():
    """Read HEARTLAND_TOKEN / HEARTLAND_SUBDOMAIN from a .env file next to
    this script, without requiring python-dotenv."""
    env_path = Path(__file__).resolve().parent / ".env"
    if env_path.exists():
        for line in env_path.read_text().splitlines():
            if line.strip() and not line.startswith("#") and "=" in line:
                k, v = line.split("=", 1)
                os.environ.setdefault(k.strip(), v.strip())

_load_heartland_env()
HEARTLAND_TOKEN = os.environ.get("HEARTLAND_TOKEN")
HEARTLAND_SUBDOMAIN = os.environ.get("HEARTLAND_SUBDOMAIN")
HEARTLAND_BASE_URL = (f"https://{HEARTLAND_SUBDOMAIN}.retail.heartland.us/api"
                      if HEARTLAND_SUBDOMAIN else None)
HEARTLAND_HEADERS = {"Authorization": f"Bearer {HEARTLAND_TOKEN}"} if HEARTLAND_TOKEN else {}

try:
    from vendor_id_map import VENDOR_ID_MAP   # generated by api_test_6.py
except ImportError:
    VENDOR_ID_MAP = {}
    print("  WARNING: vendor_id_map.py not found next to this script; "
          "Heartland live dedupe will be skipped (grids will always be "
          "generated, never checked against existing ones).")

_VENDOR_ID_LOOKUP = {}

def _rebuild_vendor_lookups(name_to_id):
    """Point both vendor lookups at one name->id mapping. _VENDOR_LOOKUP
    resolves the exact Heartland spelling for the import file; _VENDOR_ID_LOOKUP
    resolves the internal id for API calls. Building them from the same source
    is what keeps them from drifting apart."""
    _VENDOR_LOOKUP.clear()
    _VENDOR_ID_LOOKUP.clear()
    for name, vid in name_to_id.items():
        key = _vendor_key(name)
        _VENDOR_LOOKUP.setdefault(key, name)
        _VENDOR_ID_LOOKUP.setdefault(key, vid)

# Offline baseline: the committed snapshot, plus any name in KNOWN_VENDOR_NAMES
# that predates it. Replaced wholesale by the live list on every run (see
# refresh_vendor_lists) -- this is only what we fall back to when Heartland is
# unreachable, so it is never written to and is allowed to age.
_rebuild_vendor_lookups(VENDOR_ID_MAP)
for _name in KNOWN_VENDOR_NAMES:
    _VENDOR_LOOKUP.setdefault(_vendor_key(_name), _name)

def resolve_vendor_id(normalized_vendor_name: str):
    """Map an already-normalize_vendor()'d name to its Heartland internal
    vendor id. Returns None if unmapped (e.g. a vendor added in Heartland
    since the last refresh, when the refresh couldn't reach the API)."""
    return _VENDOR_ID_LOOKUP.get(_vendor_key(normalized_vendor_name))

def _heartland_fetch_vendors():
    """Every vendor in Heartland -> {name: id}, or None if the fetch failed.
    None means 'keep using the snapshot', never 'there are no vendors'."""
    if not (HEARTLAND_TOKEN and HEARTLAND_BASE_URL):
        return None
    out, page = {}, 1
    while True:
        try:
            resp = requests.get(f"{HEARTLAND_BASE_URL}/purchasing/vendors",
                                headers=HEARTLAND_HEADERS,
                                params={"per_page": 100, "page": page}, timeout=15)
        except requests.RequestException:
            return None
        if resp.status_code != 200:
            return None
        data = resp.json()
        for v in data.get("results", []):
            name, vid = (v.get("name") or "").strip(), v.get("id")
            if name and vid is not None:
                out[name] = vid
        if page >= data.get("pages", 1):
            break
        page += 1
    return out or None

VENDOR_REFRESH_SECONDS = 24 * 60 * 60    # once a day is plenty; vendors are
                                         # added weekly, not hourly
_VENDORS_REFRESHED_AT = 0.0

def refresh_vendor_lists(force=False):
    """Pull the current vendor list from Heartland, at most once a day.

    New vendors are added in Heartland every week or so, and a stale list is
    expensive in both directions: an unknown name falls through to a guessed
    spelling, which creates a DUPLICATE vendor record on import, and an
    unmapped id silently disables the dedupe check. Refreshing live costs one
    paginated fetch (~3s for ~450 vendors) against a run that already spends
    ten seconds on per-grid dedupe calls.

    Held in memory only -- vendor_id_map.py is deliberately never rewritten, so
    every machine running this converter reads the live list straight from the
    API and no clone ever has to modify (or commit) the file. That also keeps
    the shop machine git-clean, so Pull_Updates.command can't hit a merge
    conflict. On any failure the committed snapshot stands.

    The interval matters because converter_app.py is a long-running server: a
    once-per-process flag would refresh at first use and then go stale for as
    long as the window stays open, which is days.
    """
    global _VENDORS_REFRESHED_AT
    if not force and (time.time() - _VENDORS_REFRESHED_AT) < VENDOR_REFRESH_SECONDS:
        return
    live = _heartland_fetch_vendors()
    if not live:
        # Deliberately do NOT stamp the clock here: a dropped wifi connection
        # on the first run of the day would otherwise pin the stale list for a
        # full 24h even after the network came back. Retry on the next run.
        print("  NOTE: couldn't refresh the vendor list from Heartland (offline, or "
              "the token in .env is bad). Using the bundled list, which may be "
              "missing vendors added recently -- watch for 'vendor not found' warnings.")
        return
    _VENDORS_REFRESHED_AT = time.time()
    added = sorted(set(live) - set(VENDOR_ID_MAP))
    _rebuild_vendor_lookups(live)
    print(f"  refreshed vendor list from Heartland ({len(live)} vendors)")
    if added:
        print(f"  NOTE: {len(added)} vendor(s) new since the bundled list was built: "
              f"{', '.join(added[:8])}{' ...' if len(added) > 8 else ''}")

# Dedupe deliberately asks /api/items, NOT /api/item_grids, even though the
# question is "does this grid exist". Two measured reasons:
#
#   1. The grid-level 'item_custom@color' field is blank on virtually every
#      grid this converter creates -- 0 of the 25 most recently imported
#      grids carry one, and only 6 of 411 sampled grids created since 2025.
#      A query filtering on a real color name therefore never matched, so
#      dedupe silently returned False for everything and every run re-emitted
#      grids that already existed.
#   2. Heartland's custom-field filters are case-sensitive exact matches
#      ('BLACK' hits, 'Black' returns 0), and the stored casing is
#      inconsistent even within one style (CJ3548 has both 'BLACK' and
#      'Charcoal'). Filtering on color server-side can't be made reliable.
#
# Items do carry color, and heartland_vendor_item_index() already uppercases
# style and color on both sides, so matching there is case-insensitive by
# construction. Dropping the color from a grid query instead was rejected:
# four separate RCK8232 grids were imported on 2026-08-25, one per color, all
# with blank grid-level color -- a style-only match would have skipped three
# real colors, and a missing item is worse than a duplicate grid.
_STYLE_COLOR_CACHE = {}

def _vendor_style_colors(vendor_id):
    """{STYLE: {COLOR_KEY, ...}} for every item Heartland has for this vendor,
    derived from the same index build_po uses. Colors go through
    _color_match_key so separator style and casing don't matter. Keyed by
    style rather than flat (style, color) pairs so build_grids can answer both
    'does this exact grid exist' and 'does this style already exist in some
    OTHER colour spelling' without a second pass. None if the fetch failed, so
    callers can tell 'no match' apart from 'couldn't check'."""
    if vendor_id in _STYLE_COLOR_CACHE:
        return _STYLE_COLOR_CACHE[vendor_id]
    index = heartland_vendor_item_index(vendor_id)
    if index is None:
        return None
    by_style = {}
    for (style, color, _size) in index:
        by_style.setdefault(style, set()).add(color)
    _STYLE_COLOR_CACHE[vendor_id] = by_style
    return by_style

def dedupe_index_for_vendor(vendor, vendor_id):
    """The style->colors map used to skip grids Heartland already has, and a
    printed one-line statement of whether that check is actually live.

    The status line is the point. build_grids treats "couldn't check" as
    "emit the grid" (a visible duplicate the reviewer can delete beats a
    silently missing item), but two of the three ways to get there used to
    print nothing at all -- a vendor name Heartland doesn't recognise, and a
    missing .env -- so a run with the check entirely disabled was
    indistinguishable in the log from a run where every grid was genuinely
    new. That is how duplicate grids reached Heartland unnoticed.

    Note a grid imported with Generate Items = false (unknown size
    convention) has no items, so it reads as absent here and is re-emitted
    every run until its items exist. Also deliberate, same reasoning."""
    if not (HEARTLAND_TOKEN and HEARTLAND_BASE_URL):
        print(f"  WARNING: duplicate check OFF for '{vendor}' -- no Heartland "
              f"credentials found in .env. Every grid for this vendor is "
              f"emitted unchecked; look for duplicates before importing.")
        return None
    if not vendor_id:
        print(f"  WARNING: duplicate check OFF for '{vendor}' -- that is not a "
              f"known Heartland vendor, so its existing items can't be looked "
              f"up. Every grid for this vendor is emitted unchecked. Fix the "
              f"vendor name (Vendor override box, or VENDOR_MAP) and re-run to "
              f"check properly.")
        return None
    by_style = _vendor_style_colors(vendor_id)
    if by_style is None:
        print(f"  WARNING: duplicate check OFF for '{vendor}' -- couldn't reach "
              f"Heartland to list its items. Every grid for this vendor is "
              f"emitted unchecked; verify manually before importing.")
        return None
    combos = sum(len(v) for v in by_style.values())
    print(f"  duplicate check ON for '{vendor}' -- {len(by_style)} style(s) / "
          f"{combos} style-colour combination(s) already in Heartland")
    return by_style

# Which /api/items field carries the "Item #" Heartland's PO import expects.
# Vendors proved id != public_id (the UI's "#" column was public_id), so this
# is deliberately a single switch: flip to "public_id" if api_test_8 shows the
# UI value matches public_id instead of id.
ITEM_NUMBER_FIELD = "id"

def _item_number(it):
    v = it.get(ITEM_NUMBER_FIELD)
    return str(v).strip() if v not in (None, "") else ""

def _item_custom(it):
    """Items nest custom fields under 'custom'; item_grids uses 'item_custom'.
    Accept either so a schema surprise doesn't silently produce an empty index."""
    for k in ("custom", "item_custom"):
        v = it.get(k)
        if isinstance(v, dict):
            return v
    return {}

_ITEM_INDEX_CACHE = {}

def clear_item_index_cache():
    """Drop cached vendor item indexes so the next lookup refetches. Called at
    the start of each build_po run: within a run the cache saves one request
    per JOOR line, but between runs new items may have been generated in
    Heartland (e.g. by the grids import that just happened), and a stale index
    would silently push those lines into the 'missing' warning. The same
    applies to build_grids' dedupe, which reads the same index: converter_app
    is a long-running server, so a cache left over from an earlier run would
    miss the items the user has imported since and re-emit those grids."""
    _ITEM_INDEX_CACHE.clear()
    _STYLE_COLOR_CACHE.clear()

def heartland_vendor_item_index(vendor_id, refresh=False):
    """One paginated fetch of every item for a vendor ->
    {(STYLE, COLOR, SIZE): item_no}. Replaces the per-JOOR-line lookup (one
    HTTP request per line) with ~one request per 100 items. Returns None if
    the fetch couldn't be completed, so build_po can stop rather than emit a
    PO that's silently missing lines."""
    if not (HEARTLAND_TOKEN and HEARTLAND_BASE_URL and vendor_id):
        return None
    if not refresh and vendor_id in _ITEM_INDEX_CACHE:
        return _ITEM_INDEX_CACHE[vendor_id]
    index, page = {}, 1
    while True:
        try:
            resp = requests.get(
                f"{HEARTLAND_BASE_URL}/items",
                headers=HEARTLAND_HEADERS,
                params={"~[primary_vendor_id]": vendor_id,
                        "per_page": 100, "page": page},
                timeout=30,
            )
        except requests.RequestException as e:
            print(f"  ERROR: Heartland item fetch failed on page {page} ({e}).")
            return None
        if resp.status_code != 200:
            print(f"  ERROR: Heartland item fetch failed on page {page} "
                  f"(HTTP {resp.status_code}).")
            return None
        data = resp.json()
        for it in data.get("results", []):
            num = _item_number(it)
            if not num:
                continue
            cust = _item_custom(it)
            style = str(cust.get("style_number") or "").strip().upper()
            color = _color_match_key(cust.get("color") or "")
            size = canonical_size_token(cust.get("size") or "").upper()
            if style and color:
                index[(style, color, size)] = num
        pages = data.get("pages", 1)
        print(f"  fetched item page {page}/{pages} ({len(index)} indexed)")
        if page >= pages:
            break
        page += 1
    _ITEM_INDEX_CACHE[vendor_id] = index
    return index

def heartland_find_item(vendor_id, style_number, color, size):
    """Live-look up the Heartland Item # for a vendor + style number + color +
    size via /api/items (NOTE: this endpoint uses plain 'custom@' and
    'primary_vendor_id', unlike item_grids' 'item_custom@'/'item_primary_
    vendor_id' -- confirmed via api_test_7.py). Returns the item id as a
    string if found, else None (missing credentials, network error, unmapped
    vendor, or genuinely no match) -- callers treat None as 'no match'."""
    if not (HEARTLAND_TOKEN and HEARTLAND_BASE_URL and vendor_id):
        return None
    try:
        resp = requests.get(
            f"{HEARTLAND_BASE_URL}/items",
            headers=HEARTLAND_HEADERS,
            params={
                "~[primary_vendor_id]": vendor_id,
                "~[custom@style_number]": style_number,
                "~[custom@color]": color,
                "~[custom@size]": size,
                "per_page": 1,
            },
            timeout=10,
        )
        if resp.status_code != 200:
            print(f"  WARNING: Heartland item lookup failed (HTTP {resp.status_code}) "
                  f"for '{style_number}' / '{color}' / '{size}'.")
            return None
        results = resp.json().get("results", [])
        return str(results[0]["id"]) if results else None
    except requests.RequestException as e:
        print(f"  WARNING: Heartland item lookup errored ({e}) for "
              f"'{style_number}' / '{color}' / '{size}'.")
        return None

def color_code(color: str) -> str:
    words = re.split(r"[\s\-/]+", color.strip())
    return "".join(w[0] for w in words if w and w[0].isalnum()).upper()

def slug(s: str) -> str:
    return re.sub(r"[^A-Z0-9]+", "-", s.upper()).strip("-")

def _num(v):
    """'1' / 1 / 1.0 / '' -> float or None."""
    if v is None:
        return None
    s = str(v).strip()
    if not s:
        return None
    try:
        return float(s)
    except ValueError:
        return None

# ------------------------- size completion engine ---------------------------
LETTER_LADDER = ["XXXS","XXS","XS","S","M","L","XL","XXL","XXXL","XXXXL"]
LETTER_BASE = ("XS","L")

# Petite is its own ladder, NOT an extra rung on LETTER_LADDER. 98 grids in
# Heartland carry the set {P,S,M,L} (85 of them Ulla Johnson) and none of them
# carry XS, so folding P into the XS-L base run would silently add a P size to
# every ordinary letter grid the converter makes. Detection is gated on P
# actually being ordered, so plain letter orders take the branch below as before.
LETTER_PETITE_LADDER = ["P","S","M","L","XL","XXL","XXXL"]
LETTER_PETITE_BASE = ("P","L")

# Combined-letter sizing (TWP, A Shirt Thing): one item covers two rungs, so
# 'XS/S' is a single size token and NOT a color-style slash pair. 51 grids in
# Heartland use it, mostly as {P, XS/S, M/L}; without this ladder they land in
# 'unknown' and import with Generate Items = false, i.e. an empty grid.
# Only the rungs the store actually uses: adding a plausible-looking 'S/M' or
# 'L/XL' between them would insert sizes into all 51 existing grids' shape,
# and every one of those grids carries P, which is why P is the base floor.
COMBO_LADDER = ["P","XS/S","M/L","1XL"]
COMBO_BASE = ("P","M/L")

OS_TOKENS = {"OS","O/S","ONE SIZE","ONESIZE","NA", "N/A", "N/a", "Na", "na", "n/a"}

# Ring ladder: 6-8 is the store's base run, 5 and 9 only when actually ordered.
# Same semantics as LETTER_LADDER/LETTER_BASE.
RING_LADDER = ["5","6","7","8","9"]
RING_BASE = ("6","8")

# Ring sizes are indistinguishable from apparel sizes by token alone -- an
# all-even ring order (6,8) would otherwise match the numeric-even branch and
# generate a 00-20 apparel ladder. The style name is the only signal that
# travels with every export format, so detection keys off it. 's?' catches the
# plural; the word boundary keeps 'Earring(s)', 'Herringbone', and 'Ringer'
# from matching, which was verified against a real Jenny Bird order.
_RING_NAME_RE = re.compile(r"(?i)\brings?\b")

def looks_like_ring(style_name) -> bool:
    """True when a line's style name says it's a ring. Deliberately narrow:
    a false positive puts apparel on a 5-9 ladder, so prefer missing an
    oddly-named ring (it falls back to today's behavior) over claiming one."""
    return bool(_RING_NAME_RE.search(str(style_name or "")))

def consolidate_size_suffixes(rows):
    """Some vendors give every size its own style number (Jenny Bird's Renzo
    Ring is 'JB4014-6-HPG-W' in a 6 and 'JB4014-7-HPG-W' in a 7). Left alone
    those become one grid per size, which duplicates a grid Heartland may
    already hold in consolidated form. Strip the size segment so the sizes
    collapse into a single grid keyed on the base style ('JB4014-HPG-W').

    Two guards, both derived from what the live catalog actually contains:

    'OS' is never stripped. A one-size item has no siblings to merge with, and
    trailing '-OS' is common enough (Staud) that stripping it would rewrite a
    style number that already matches items in Heartland and break the PO
    lookup -- it was 7 of the 9 false positives in a 5,500-item sample.

    A family is only consolidated when 2+ distinct sizes share the same base
    and color. A lone '-S-' or '-6' segment is far more likely to be part of
    the vendor's own numbering than a size axis, and rewriting it would
    orphan the item for no gain. Every non-Jenny-Bird match in that sample
    was a single-size family.
    """
    fams, marks = {}, {}
    for i, r in enumerate(rows):
        sn = str(r["style_number"]).strip()
        size = canonical_size_token(r["size"])
        if size == "OS" or not sn:
            continue
        parts = sn.split("-")
        if len(parts) < 2 or size not in parts:
            continue
        base = "-".join(p for p in parts if p != size)
        if not base:
            continue                      # style was nothing but the size
        key = (str(r["seller_account"]).strip(), base, normalize_color(r["color"]))
        fams.setdefault(key, set()).add(size)
        marks[i] = (key, base, sn)

    out, renamed = [], {}
    for i, r in enumerate(rows):
        if i in marks:
            key, base, sn = marks[i]
            if len(fams[key]) > 1:        # a real size axis, not a one-off segment
                nr = dict(r); nr["style_number"] = base
                renamed.setdefault(sn, base)
                out.append(nr)
                continue
        out.append(r)
    for old, new in sorted(renamed.items()):
        print(f"  NOTE: '{old}' carries its size in the style number; consolidating "
              f"into '{new}' so every size lands in one grid.")
    return out

def _size_encoded_in_style_number(style_number, sizes) -> bool:
    """True when the vendor puts the ring size in the style number itself, so
    each size is its own style (Jenny Bird: 'JB4014-6-HPG-W' is the size 6,
    'JB4014-7-HPG-W' the size 7). build_grids groups by style number, so these
    are separate grids and completing each to the full 6-8 base run would
    create the same physical ring three times over, once per grid."""
    parts = re.split(r"[-_/ ]+", str(style_number).strip())
    return any(canonical_size_token(s) in parts for s in sizes)

# Apparel numeric-even ladder (dresses/tops/etc): '00' is the rung one step
# below size 0 (NOT the same as '0'); the run commonly extends above 10 to
# 12/14/16/18+. Kept as strings throughout so '00' stays distinct from '0'.
NUMERIC_EVEN_LADDER = ["00","0","2","4","6","8","10","12","14","16","18","20"]

def _consecutive_numeric_key(tok):
    """Sort key for a single-step numeric ladder. '00' is the rung one below
    '0' (see NUMERIC_EVEN_LADDER's comment), which plain int() destroys."""
    return -1 if str(tok).strip() == "00" else int(tok)

def _is_consecutive_numeric_token(s):
    """A whole non-negative number, or the distinct '00' rung."""
    s = str(s).strip()
    if s == "00":
        return True
    try:
        return int(s) >= 0
    except ValueError:
        return False

def _is_numeric_even_token(s):
    s = str(s).strip()
    if s == "00":
        return True
    try:
        v = int(s)
    except ValueError:
        return False
    return v >= 0 and v % 2 == 0

def _numeric_even_ladder_upto(sizes):
    """Extend the base 00..20 ladder further if an order runs above it."""
    ladder = list(NUMERIC_EVEN_LADDER)
    hi_val = max(-1 if str(s).strip() == "00" else int(s) for s in sizes)
    last_val = int(ladder[-1])
    while last_val < hi_val:
        last_val += 2
        ladder.append(str(last_val))
    return ladder

def _size_sort_key(size):
    """Retail order for a loose pile of size tokens, used wherever no ladder
    applies (the 'unknown' fallback and the ring-in-style-number branch).
    Letter ladders first in ladder order, then numbers, then anything else
    alphabetically. A plain alphabetical sort is what turned Ulla Johnson's
    P,S,M into 'M,P,S' in Heartland."""
    s = str(size).strip().upper()
    if s in LETTER_LADDER:
        return (0, LETTER_LADDER.index(s), "")
    if s == "P":
        # Petite has no rung on LETTER_LADDER (every other LETTER_PETITE_LADDER
        # token does, and is caught above). Sort it just below S so a mixed
        # pile still reads P,S,M,L rather than alphabetically.
        return (0, LETTER_LADDER.index("S") - 0.5, "")
    if s in COMBO_LADDER:
        # 'XS/S' etc: sort by the rung it starts on, so a pile that stays
        # 'unknown' still reads P, XS/S, M/L rather than alphabetically.
        return (0, LETTER_LADDER.index("S") - 0.5 + COMBO_LADDER.index(s) * 0.01, "")
    if s == "00":
        return (1, -1, "")           # '00' is the rung below '0', not zero
    try:
        return (1, float(s), "")
    except ValueError:
        return (2, 0, s)

def detect_convention(sizes, is_ring=False):
    """Return (name, ladder, base_lo_idx, base_hi_idx) or fallback.
    is_ring comes from the style name (see looks_like_ring) because ring and
    apparel size tokens are identical -- without it, '6,8' is ambiguous."""
    up = [s.upper() for s in sizes]
    if all(s in OS_TOKENS for s in up):
        return ("os", ["OS"], 0, 0)
    # The '"P" in up' gate is what keeps ordinary XS-L orders on LETTER_LADDER.
    # Without it -- i.e. if P were simply added as a rung there -- every letter
    # grid the converter makes would grow a Petite size, since LETTER_BASE
    # spans XS to L and completion fills the whole run.
    if "P" in up and all(s in LETTER_PETITE_LADDER for s in up):
        return ("letters-petite", LETTER_PETITE_LADDER,
                LETTER_PETITE_LADDER.index(LETTER_PETITE_BASE[0]),
                LETTER_PETITE_LADDER.index(LETTER_PETITE_BASE[1]))
    if all(s in LETTER_LADDER for s in up):
        return ("letters", LETTER_LADDER,
                LETTER_LADDER.index(LETTER_BASE[0]), LETTER_LADDER.index(LETTER_BASE[1]))
    # Needs a slashed token present, so a plain {P} or {P,S} order stays on the
    # petite ladder above rather than being read as combined sizing.
    if any("/" in s for s in up) and all(s in COMBO_LADDER for s in up):
        return ("combo-letters", COMBO_LADDER,
                COMBO_LADDER.index(COMBO_BASE[0]), COMBO_LADDER.index(COMBO_BASE[1]))
    # MUST precede the numeric branches: numeric-even would claim an all-even
    # ring order (6,8) and hand back a 00-20 apparel ladder.
    if is_ring and all(canonical_size_token(s) in RING_LADDER for s in sizes):
        return ("ring", RING_LADDER,
                RING_LADDER.index(RING_BASE[0]), RING_LADDER.index(RING_BASE[1]))
    if all(_is_numeric_even_token(s) for s in sizes):
        # Handles plain runs (2,4,6,8), sizes that dip to '00', and runs that
        # climb past 10 (12,14...) -- all as one continuous even ladder.
        ladder = _numeric_even_ladder_upto(sizes)
        return ("numeric-even", ladder, ladder.index("0"), ladder.index("10"))
    if all(re.fullmatch(r"\d{1,2}-\d{1,2}", s) for s in sizes):
        # Literal hyphenated size tokens (e.g. Veronica Beard '00-14') are a
        # single named size, not a numeric run -- treat as their own closed
        # ladder so complete_sizes doesn't warn "unrecognized convention".
        ladder = sorted(set(sizes))
        return ("literal-hyphenated", ladder, 0, len(ladder) - 1)
    # Brands on a small consecutive ladder (Lisa Yang: 00,0,1,2,3) have no
    # store-standard base run, so there is nothing to complete TO. The old
    # numeric-0-4 branch guessed one and invented sizes: an order of 0,1,00
    # came back from Heartland as 0,1,2,3,4 -- three sizes never ordered, and
    # the '00' rung lost because int("00") == 0 collapsed it into '0'.
    # Emit exactly what was ordered instead, in numeric order, '00' intact.
    # numeric-even is checked first, so plain even runs still complete to 0-10.
    if all(_is_consecutive_numeric_token(s) for s in sizes):
        toks = {canonical_size_token(s) for s in sizes}
        if max(_consecutive_numeric_key(t) for t in toks) <= 12:
            ladder = sorted(toks, key=_consecutive_numeric_key)
            return ("numeric-ordered", ladder, 0, len(ladder) - 1)
    try:
        nums = sorted({int(s) for s in sizes})
    except ValueError:
        return ("unknown", None, None, None)
    lo, hi = nums[0], nums[-1]
    if hi <= 12:
        # Only negative sizes reach here now; every non-negative all-numeric
        # set was claimed by numeric-even or numeric-ordered above.
        return ("unknown", None, None, None)
    if 20 <= lo and hi <= 45:                      # denim 24-29 base, step 1
        ladder = list(range(min(lo, 24), max(hi, 29) + 1))
        return ("denim-24-29", ladder, ladder.index(24), ladder.index(29))
    if lo >= 60 and all(n % 5 == 0 for n in nums): # euro 70-90 base, step 5
        ladder = list(range(min(lo, 70), max(hi, 90) + 5, 5))
        return ("euro-70-90", ladder, ladder.index(70), ladder.index(90))
    return ("unknown", None, None, None)

def complete_sizes(ordered_sizes, is_ring=False):
    """Continuous ladder span from min(base, ordered) to max(base, ordered)."""
    name, ladder, b_lo, b_hi = detect_convention(ordered_sizes, is_ring)
    if name == "os":
        return ["OS"], name
    if name == "unknown":
        print(f"  WARNING: unrecognized size convention {sorted(set(ordered_sizes))}; "
              f"using ordered sizes as-is (no completion).")
        # sorted() alone is alphabetical, which is what put Ulla Johnson's
        # P,S,M into Heartland as 'M,P,S'. _size_sort_key keeps retail order
        # even for sets no ladder recognizes.
        return sorted({str(s) for s in ordered_sizes}, key=_size_sort_key), name
    if name in ("letters", "letters-petite", "combo-letters"):
        idxs = [ladder.index(s.upper()) for s in ordered_sizes]
    elif name == "numeric-even":
        idxs = [ladder.index("00" if str(s).strip() == "00" else str(int(s)))
                for s in ordered_sizes]
    elif name == "literal-hyphenated":
        # ladder is the literal string tokens themselves ('00-14'); never int()
        idxs = [ladder.index(str(s).strip()) for s in ordered_sizes]
    elif name == "numeric-ordered":
        # ladder IS the ordered set, so the span below covers all of it and
        # nothing is added -- '00' must stay a string here too.
        idxs = [ladder.index(canonical_size_token(s)) for s in ordered_sizes]
    elif name == "ring":
        # string ladder like letters, but the tokens are digits -- go through
        # canonical_size_token so '07'/'SIZE 7' land on the same rung as '7'
        idxs = [ladder.index(canonical_size_token(s)) for s in ordered_sizes]
    else:
        idxs = [ladder.index(int(s)) for s in ordered_sizes]
    lo, hi = min(min(idxs), b_lo), max(max(idxs), b_hi)
    return [str(ladder[i]) for i in range(lo, hi + 1)], name

def canonical_size_token(raw) -> str:
    """Normalize a single raw JOOR size token to the exact string Heartland
    stores as an item's size after grid generation (same rules build_grids
    applies via detect_convention/complete_sizes, but for one token instead
    of filling a range) -- used by build_po so its live item lookup always
    queries with the same token build_grids used to create the item."""
    s = clean_size(raw)
    su = s.upper()
    if su in OS_TOKENS:
        return "OS"
    if su in LETTER_LADDER:
        return su
    if su == "00":
        return "00"
    try:
        return str(int(su))
    except ValueError:
        return s   # unrecognized convention: build_grids keeps original casing too

def _numeric_even_span(lo_token, hi_token):
    """Inclusive even-step run between two numeric-even tokens, e.g. '00'..'14'."""
    ladder = _numeric_even_ladder_upto([lo_token, hi_token])
    lo_key = "00" if lo_token == "00" else str(int(lo_token))
    hi_key = "00" if hi_token == "00" else str(int(hi_token))
    return ladder[ladder.index(lo_key):ladder.index(hi_key) + 1]

def expand_size_ranges(rows):
    """Sizes like '00-14' (e.g. Veronica Beard) are a single literal size
    token, NOT a range meaning 'one of each size in between'. Just clean and
    pass through -- no expansion."""
    out = []
    for r in rows:
        nr = dict(r); nr["size"] = clean_size(r.get("size", ""))
        out.append(nr)
    return out

# ------------------------- pricing -------------------------------------------
def ceil_dollars(val) -> int:
    return int(Decimal(str(val)).to_integral_value(rounding=ROUND_CEILING))

def retail_price(unit_cost, sugg_retail=None) -> int:
    """JOOR Sugg. Retail if present, else MARGIN x cost; always ceil to whole $."""
    if sugg_retail:
        return ceil_dollars(sugg_retail)
    return ceil_dollars(Decimal(str(unit_cost)) * MARGIN)

# ------------------------- io: format detection & parsing --------------------
def _load_sheet_rows(path):
    """xlsx -> (worksheet, list-of-list of string cell values). data_only skips formulas."""
    from openpyxl import load_workbook
    wb = load_workbook(path, data_only=True)
    ws = wb.active
    rows = [[("" if c is None else str(c).strip()) for c in r]
            for r in ws.iter_rows(values_only=True)]
    return ws, rows

def _find_meta(rows, label):
    """Find 'Label:' either as its own cell (value = next non-empty cell to the
    right) or embedded 'Label: value' inside one cell."""
    lab = label.lower().rstrip(":")
    for r in rows:
        for i, cell in enumerate(r):
            low = cell.lower()
            if low.rstrip(":") == lab and cell.endswith(":"):
                for nxt in r[i+1:]:
                    if nxt:
                        return nxt
            m = re.match(rf"{re.escape(lab)}\s*:\s*(.+)", low)
            if m:
                return cell.split(":", 1)[1].strip()
    return ""

def _looks_like_color_code(s: str) -> bool:
    """True for short alphanumeric codes like 'C0308' or 'R0272' (no spaces,
    mixes letters+digits, short) as opposed to a real color name like
    'Oatmeal Melange Multi' or 'Cardinal Red'."""
    s = s.strip()
    if not s:
        return True                     # empty -> not a usable name either
    if " " in s or "/" in s:
        return False                    # multi-word or slash-combo -> real name
    return bool(re.fullmatch(r"[A-Za-z]{0,3}\d+[A-Za-z0-9]*", s))

def resolve_color(color_val, color_code_val):
    """JOOR's 'Color' vs 'Color Code' columns hold the descriptive name and the
    internal code in either order depending on the vendor/brand (confirmed:
    DOEN puts the name in 'Color'; Veronica Beard puts it in 'Color Code').
    Pick whichever cell actually looks like a name; fall back to 'Color' if
    both or neither look like one."""
    color_val, color_code_val = (color_val or "").strip(), (color_code_val or "").strip()
    if color_val and not _looks_like_color_code(color_val):
        return color_val
    if color_code_val and not _looks_like_color_code(color_code_val):
        return color_code_val
    return color_val or color_code_val  # both look like codes (or one's blank): best guess

# ------------------------- NuOrder support -----------------------------------
# NuOrder's "Order Data" export is already long (one row per style/color/size),
# so it needs a reader, not a new pipeline. Two wrinkles: its column set varies
# between exports (some carry SKU/Long Style Number/ship dates, others carry UPC
# and no dates), so columns are resolved by header NAME and all but a few are
# optional; and the vendor is never in the file at all, so the caller supplies
# it. Grids only -- NuOrder POs are built in Heartland by selecting the grids,
# so no order number or ship dates are needed here.
NUORDER_REQUIRED = ("style number", "name", "color", "size", "qty")

def looks_like_nuorder(header_row) -> bool:
    h = {str(c).strip().lower() for c in header_row}
    return "_id" in h and "units per pack" in h

def read_nuorder(path, vendor):
    """NuOrder 'Order Data' xlsx -> the same long row dicts build_grids consumes."""
    from openpyxl import load_workbook
    wb = load_workbook(path, data_only=True)
    ws = next((w for w in wb.worksheets if "order data" in w.title.lower()),
              wb.worksheets[0])          # the "Summary" tab is informational only
    rows = [[("" if c is None else str(c).strip()) for c in r]
            for r in ws.iter_rows(values_only=True)]
    if not rows:
        sys.exit("ERROR: the NuOrder file appears to be empty.")
    col = {c.lower(): i for i, c in enumerate(rows[0]) if c}
    missing = [c for c in NUORDER_REQUIRED if c not in col]
    if missing:
        sys.exit(f"ERROR: this doesn't look like a NuOrder order export "
                 f"(missing column(s): {', '.join(missing)}).")
    if not str(vendor).strip():
        sys.exit("ERROR: no vendor given. NuOrder doesn't put the brand name in "
                 "the export, so it has to be typed in.")

    def cell(r, name):
        i = col.get(name)
        return r[i] if i is not None and i < len(r) else ""

    long_rows, no_qty = [], 0
    for r in rows[1:]:
        style_no = cell(r, "style number")
        if not style_no:
            continue                                    # blank / footer row
        qty = _num(cell(r, "qty"))
        if not qty:
            no_qty += 1        # NuOrder lists the vendor's whole size run; rows
            continue           # with no qty were offered, not ordered
        long_rows.append({
            "seller_account": vendor,
            "style_number":   style_no,
            "style_name":     strip_style_number_prefix(cell(r, "name"), style_no),
            "color":          cell(r, "color"),
            "size":           clean_size(cell(r, "size")),
            "quantity":       qty * (_num(cell(r, "units per pack")) or 1),
            "item_price":     _num(cell(r, "wholesale (usd)")) or 0,
            "sugg_retail":    _num(cell(r, "m.s.r.p (usd)")),
            "order_number":   "",
            SEASON_FROM:      "",
        })
    print(f"  [nuorder] vendor='{vendor}'  {len(long_rows)} ordered line(s); "
          f"skipped {no_qty} zero-qty size row(s)")
    if not long_rows:
        sys.exit("ERROR: no ordered lines found (every row had a blank Qty).")
    return long_rows

def strip_style_number_prefix(style_name: str, style_number: str) -> str:
    """Some vendors (e.g. AMO) prefix Style Name with the style number in
    JOOR's Order Summary export: 'WFB270-1450C - EASY ARMY SLOUCH CROP'.
    Drop it when it matches the row's own Style Number, so Description is the
    product name alone. Only strips an exact match -- a name that merely
    starts with digits is left intact."""
    name, num = str(style_name).strip(), str(style_number).strip()
    if not num:
        return name
    m = re.match(rf"^{re.escape(num)}\s*[-–—:]\s*(.+)$", name, re.IGNORECASE)
    return m.group(1).strip() if m else name
    
def parse_order_to_size(path):
    """JOOR 'Order to Size' wide xlsx -> normalized long rows."""
    _, rows = _load_sheet_rows(path)

    # --- header row of the line-item table
    hdr_idx = next((i for i, r in enumerate(rows)
                    if "style number" in [c.lower() for c in r]
                    and "style name" in [c.lower() for c in r]), None)
    if hdr_idx is None:
        sys.exit("ERROR: couldn't find the line-item header row (Style Name / Style Number).")
    hdr = rows[hdr_idx]
    col = {c.lower(): i for i, c in enumerate(hdr) if c}
    size_cols = [(i, c) for i, c in enumerate(hdr) if c
                 and c.lower() not in WIDE_NON_SIZE
                 and not c.lower().startswith(WIDE_NON_SIZE_PREFIXES)]
    ws_col = next((i for c, i in col.items() if c.startswith("wholesale")), None)
    sr_col = next((i for c, i in col.items() if c.startswith(("sugg. retail", "sugg retail"))), None)
    if ws_col is None:
        sys.exit("ERROR: no WholeSale column found in the line-item table.")

    # --- order-level metadata
    vendor_raw = next((c for r in rows[:hdr_idx] for c in r if c), "")
    if vendor_raw.endswith(":") or not vendor_raw:
        vendor_raw = ""
        print("  WARNING: couldn't identify the vendor cell; set VENDOR_MAP or check output.")
    order_no  = _find_meta(rows[:hdr_idx], "PO Number")
    start_ship = _find_meta(rows[:hdr_idx], "Start Ship")
    end_ship   = _find_meta(rows[:hdr_idx], "Complete")
    print(f"  [order-to-size] vendor='{vendor_raw}'  PO={order_no}  "
          f"ship {start_ship} -> {end_ship}  (no linesheet column: Season will be blank)")

    # --- melt wide line rows into the same long dicts the pipeline expects
    long_rows, line_row_idx = [], {}
    for i in range(hdr_idx + 1, len(rows)):
        r = rows[i]
        style_no  = r[col["style number"]] if col["style number"] < len(r) else ""
        style_name = r[col["style name"]] if col["style name"] < len(r) else ""
        if not style_no or not style_name:
            continue                                   # totals / footer rows
        clean_name = strip_style_number_prefix(style_name, style_no)
        if clean_name != style_name:
            print(f"  NOTE: '{style_name}': stripped the style number prefix from "
                  f"Style Name -> '{clean_name}'.")
            style_name = clean_name                                 # totals / footer rows
        color_raw  = r[col["color"]] if col["color"] < len(r) else ""
        code_col_idx = col.get("color code")
        color_code_raw = r[code_col_idx] if code_col_idx is not None and code_col_idx < len(r) else ""
        color = resolve_color(color_raw, color_code_raw)
        if color and color != color_raw.strip():
            print(f"  NOTE: '{style_name}': using Color Code column value "
                  f"'{color}' as the real color name ('{color_raw}' in the "
                  f"Color column looks like an internal code for this vendor).")
        line_row_idx[i] = (style_no, color)            # for image anchoring
        for ci, raw_label in size_cols:
            qty = _num(r[ci]) if ci < len(r) else None
            if not qty:
                continue
            size_label = clean_size(raw_label)
            run_sizes = [size_label]
            for size_val in run_sizes:
                long_rows.append({
                    "seller_account": vendor_raw,
                    "style_number":  style_no,
                    "style_name":    style_name,
                    "color":         color,
                    "size":          size_val,
                    "quantity":      qty,
                    "item_price":    _num(r[ws_col]) or 0,
                    "sugg_retail":   _num(r[sr_col]) if sr_col is not None else None,
                    "order_number":  order_no,
                    "start_date":    start_ship,
                    "cancel_date":   end_ship,
                    SEASON_FROM:     "",
                })
    return long_rows

def read_joor(path):
    """Read a JOOR export in either format -> list of long row dicts.
    Auto-detects: Basic Export (long) vs Order-to-Size (wide, formatted xlsx).
    Both branches return the SAME shape (a plain list) -- they used to differ,
    which crashed whichever caller assumed the other."""
    p = Path(path)
    if p.suffix.lower() in (".xlsx", ".xlsm"):
        _, rows = _load_sheet_rows(p)
        first_hdr = [h.lower() for h in rows[0]] if rows else []
        if looks_like_nuorder(rows[0] if rows else []):
            sys.exit("ERROR: this is a NuOrder export, not a JOOR export.\n"
                     "  -> use the NuOrder section of the converter, which also "
                     "asks for the vendor name (NuOrder doesn't put it in the file).")
        if "seller_account" not in first_hdr:                    # wide format
            return parse_order_to_size(p)
        rows = [[c for c in r] for r in rows]
    else:
        raw = p.read_bytes()
        for enc in ("utf-8-sig", "utf-8", "cp1252", "latin-1"):
            try:
                text = raw.decode(enc); break
            except UnicodeDecodeError:
                continue
        delim = "\t" if text.splitlines()[0].count("\t") > text.splitlines()[0].count(",") else ","
        rows = list(csv.reader(text.splitlines(), delimiter=delim))
    hdr = [h.strip().lower() for h in rows[0]]
    long_rows = [dict(zip(hdr, r)) for r in rows[1:] if any(str(c).strip() for c in r)]
    return expand_size_ranges(long_rows)

def write_xlsx(path, headers, rows, text_cols=(), dropdowns=None):
    from openpyxl import Workbook
    from openpyxl.styles import Font
    from openpyxl.utils import get_column_letter
    from openpyxl.worksheet.datavalidation import DataValidation
    wb = Workbook(); ws = wb.active
    bold = Font(name="Arial", size=10, bold=True); norm = Font(name="Arial", size=10)
    for c, h in enumerate(headers, 1):
        ws.cell(row=1, column=c, value=h).font = bold
    for r_i, row in enumerate(rows, 2):
        for c, h in enumerate(headers, 1):
            v = row.get(h)
            cell = ws.cell(row=r_i, column=c)
            if h in text_cols and v is not None:
                cell.value = str(v); cell.number_format = "@"
            else:
                cell.value = v
            cell.font = norm
    for c, h in enumerate(headers, 1):
        vals = [h] + [str(r.get(h) or "") for r in rows]
        ws.column_dimensions[get_column_letter(c)].width = min(max(len(s) for s in vals) + 2, 34)
    ws.freeze_panes = "A2"

    if dropdowns:
        # Hidden helper sheet holding each option list; data validation formulas
        # reference it by range (more reliable across Excel/Numbers than long
        # inline comma lists, which also risk hitting the ~255-char limit).
        lists_ws = wb.create_sheet("Lists")
        lists_ws.sheet_state = "hidden"
        last_row = max(len(rows) + 100, 200)   # room for manually added rows later
        for col_i, (header, options) in enumerate(dropdowns.items(), 1):
            if header not in headers:
                continue
            target_col_idx = headers.index(header) + 1
            col_letter = get_column_letter(target_col_idx)  # use target, not col_i
            for r_i, opt in enumerate(options, 1):
                lists_ws.cell(row=r_i, column=target_col_idx, value=opt)
            dv = DataValidation(
                type="list",
                formula1=f"=Lists!${col_letter}$1:${col_letter}${len(options)}",
                allow_blank=True, showDropDown=False)
            ws.add_data_validation(dv)
            dv.add(f"{col_letter}2:{col_letter}{last_row}")

    wb.save(path)

# ------------------------- phase 1: grids ------------------------------------
def build_grids(joor_rows, out_path, vendor_override=""):
    """vendor_override forces every grid onto one exact Heartland vendor name,
    for when JOOR's spelling can't be resolved automatically (the app's Vendor
    override box). Empty means resolve per row, which is the normal path."""
    refresh_vendor_lists()     # new vendors appear weekly; a stale name duplicates
    clear_item_index_cache()   # dedupe reads it; the user may have imported since
    joor_rows = consolidate_size_suffixes(joor_rows)
    override = str(vendor_override or "").strip()
    if override:
        # Worded to read correctly on both tabs: on the JOOR tab this is an
        # override of the name in the file, on the NuOrder tab it is the only
        # source of the vendor at all.
        print(f"  NOTE: all grids assigned to vendor '{override}' "
              f"(from the Vendor field), not to any name in the order file.")
    resolved = {}              # raw seller_account -> Heartland name, so the
                               # NOTE/WARNING prints once per vendor, not per row
    groups = {}                                   # (vendor, style#, color) -> data
    for r in joor_rows:
        raw_vendor = r["seller_account"]
        if raw_vendor not in resolved:
            resolved[raw_vendor] = override or normalize_vendor(raw_vendor)
        vendor = resolved[raw_vendor]
        raw_color = r["color"]
        if _looks_like_color_code(raw_color):
            print(f"  NOTE: color '{raw_color}' for '{r['style_name']}' looks like a "
                  f"vendor code, not a color name. This export format doesn't carry a "
                  f"plain-English color column -> re-download using JOOR's 'Download "
                  f"Order Summary Excel' option instead, which does.")
        key = (vendor, r["style_number"].strip(), normalize_color(raw_color))
        g = groups.setdefault(key, {"sizes": [], "cost": r["item_price"],
                                    "sugg_retail": r.get("sugg_retail"),
                                    "style_name_raw": r["style_name"],
                                    "raw_key": (r["style_number"].strip(), r["color"]),
                                    "is_ring": looks_like_ring(r["style_name"]),
                                    "season": str(r.get(SEASON_FROM, "") or "").strip()})
        g["sizes"].append(str(r["size"]).strip())

    # grid numbers with intra-run collision handling
    codes = {}
    for (vendor, style_no, color) in groups:
        code = color_code(color)
        while (vendor, style_no, code) in codes.values():
            code = slug(color)                    # fall back to full color slug
        codes[(vendor, style_no, color)] = (vendor, style_no, code)

    # One dedupe index per vendor, built before the grid loop rather than
    # inside it: _vendor_style_colors doesn't cache a failed fetch, so a
    # per-group call re-ran a full paginated request for every grid whenever
    # Heartland was unreachable.
    dedupe = {v: dedupe_index_for_vendor(v, resolve_vendor_id(v))
              for v in sorted({v for (v, _s, _c) in groups})}

    out_rows, skipped_keys, unchecked = [], set(), 0
    detail_lines, unrecognized = [], []
    for key, g in groups.items():
        vendor, style_no, color = key
        by_style = dedupe[vendor]
        style_key = str(style_no).strip().upper()
        if by_style is None:
            unchecked += 1
        else:
            existing = by_style.get(style_key, set())
            if _color_match_key(color) in existing:
                skipped_keys.add(key)
                detail_lines.append(f"  SKIP (already exists in Heartland): {style_no} / {color} ({vendor})")
                continue
            if existing:
                # The style is in Heartland but not in this colour. Usually a
                # genuinely new colourway -- but it is also exactly what a
                # colour spelled differently looks like, which is how the
                # Generation Love 'BLACK/CREAM' vs 'BLACK / CREAM' duplicate
                # got in. Cheap to print, and the reviewer can settle it.
                print(f"  NOTE: {style_no} is new in '{color}', but Heartland "
                      f"already has {style_no} in: {', '.join(sorted(existing))}. "
                      f"If one of those is the same colour spelled differently, "
                      f"delete this grid before importing.")
        desc = title_caps(fix_mojibake(g["style_name_raw"]))
        if g["is_ring"] and _size_encoded_in_style_number(style_no, g["sizes"]):
            # One style number per size: complete nothing, or every sibling
            # grid would repeat the whole 6-8 base run.
            sizes = sorted({canonical_size_token(s) for s in g["sizes"]},
                           key=_size_sort_key)
            conv = "ring (size in style number)"
            print(f"  NOTE: '{desc}' ({style_no}): this vendor puts the ring size in "
                  f"the style number, so each size is its own grid -- generating only "
                  f"the ordered size {sizes}, not the 6-8 base run.")
        else:
            sizes, conv = complete_sizes(g["sizes"], g["is_ring"])
        price = retail_price(g["cost"], g.get("sugg_retail"))
        src = "JOOR Sugg. Retail" if g.get("sugg_retail") else f"{MARGIN}x cost"
        detail_lines.append(f"  {desc} / {color}: ordered {sorted(set(g['sizes']))} -> "
                            f"[{conv}] generated {sizes}  | price ${price} ({src})")
        if conv == "unknown":
            unrecognized.append((desc, color, sorted(set(g["sizes"]))))
            detail_lines.append(f"    -> Generate Items = false for this grid "
                                f"(sizes not recognized; add items manually in Heartland).")
        row = {h: None for h in TEMPLATE_HEADERS}
        row.update({
            "Item Grid #": "-".join(codes[key]),
            "Description": desc,
            "Default Cost": float(g["cost"]),
            "Price": price,
            "Primary Vendor": vendor,
            "Style Name": f"{desc} in {color}",
            "Style Number": style_no,
            "Season": None, # We don't use seasons anymore
            "Dimension 1": "Color", "Dimension 1 Values": color,
            "Dimension 2": "Size", "Dimension 2 Values": ",".join(sizes),
            "Active": "true",
            "Generate Items": "false" if conv == "unknown" else "true",
            "Active fit guide": "3",
        })
        out_rows.append(row)
    if unrecognized:
        print(f"WARNING: {len(unrecognized)} grid(s) had size names this converter "
              f"doesn't recognize. Their items were NOT generated -- the grid is "
              f"created empty and you must add sizes/items manually in Heartland:")
        for desc, color, raw_sizes in unrecognized:
            print(f"  - {desc} / {color}: sizes {raw_sizes}")
        print("")
    for line in detail_lines:
        print(line)
    write_xlsx(out_path, TEMPLATE_HEADERS, out_rows,
               text_cols=("Item Grid #", "Style Number"), dropdowns=DROPDOWN_LISTS)
    print(f"wrote {out_path} ({len(out_rows)} grid(s))")
    print(f"  duplicate check: {len(skipped_keys)} grid(s) skipped as already in "
          f"Heartland, {len(out_rows) - unchecked} checked and new, "
          f"{unchecked} emitted UNCHECKED")
    if unchecked:
        print(f"  WARNING: {unchecked} grid(s) could not be checked against "
              f"Heartland (see the reasons above) -- review those before importing.")
    if unrecognized:
        print(f"  NOTE: the {len(unrecognized)} grid(s) with Generate Items = false "
              f"create no items, so they will show as 'new' again on the next run "
              f"until you add their sizes in Heartland.")

# ------------------------- phase 2: PO ---------------------------------------
PO_HEADERS = ["PO #", "PO Description", "PO Vendor", "PO Start Ship", "PO End Ship", "PO Received at location",
              "Item #", "PO Line Qty", "PO Line Unit Cost", "Item Default Cost"]

def po_description(order_number, shipping_fee=None):
    """PO Description convention: '<order_number>' normally, or
    'F:$<fee>/<order_number>' when a per-order shipping fee was charged."""
    if shipping_fee:
        return f"F:${ceil_dollars(shipping_fee)}/{order_number}"
    return str(order_number)
# PO # placeholder is used because it is a required field upon PO import.
# Heartland doesn't allow auto-generation of the PO # upon PO import, so this
# placeholder allows the import to happen but requires user follow up. The JOOR
# order number goes in PO Description instead, so it's still visible/searchable
# without overriding Heartland's numbering.
RECEIVED_AT_LOCATION = "romi boutique"   # Heartland requires this field on PO import; single-location store
# Helper columns are parenthesized on purpose: Heartland won't auto-map them,
# so they import as Ignored and can't overwrite anything.

def _iget(it, *keys):
    """First non-empty value among several possible header names. Heartland's
    ITEMS export uses unprefixed headers (Description, Color, Size, Style
    Number...) while its GRIDS export prefixes them with 'Item '. Accept both."""
    for k in keys:
        v = it.get(k)
        if v not in (None, ""):
            return str(v)
    return ""

def build_po(joor_rows, out_path):
    refresh_vendor_lists()     # same reason as build_grids
    # Must run here too: the items build_grids created are keyed on the
    # consolidated style number, so looking them up under the original
    # per-size number would miss every consolidated line.
    joor_rows = consolidate_size_suffixes(joor_rows)
    clear_item_index_cache()   # items may have been generated since the last run
    out_rows, missing = [], []
    resolved = {}              # raw seller_account -> Heartland name, so the
                               # NOTE/WARNING prints once per vendor, not per line
    for r in joor_rows:
        raw_vendor = r["seller_account"]
        if raw_vendor not in resolved:
            resolved[raw_vendor] = normalize_vendor(raw_vendor)
        vendor = resolved[raw_vendor]
        vendor_id = resolve_vendor_id(vendor)
        index = heartland_vendor_item_index(vendor_id)
        if index is None:
            sys.exit(f"ERROR: couldn't load Heartland items for vendor '{vendor}'.\n"
                     "  -> check the .env credentials and network, and that this "
                     "vendor exists in vendor_id_map.py. No PO file written.")
        color = normalize_color(r["color"])
        size = canonical_size_token(r["size"])
        style_key = r["style_number"].strip().upper()
        item_no = index.get((style_key, _color_match_key(color), size.upper()))
        if not item_no:
            missing.append((r["style_number"], color, r["size"])); continue
        cost = float(r["item_price"])
        out_rows.append({
            "PO #": "CHECK HEARTLAND HIGHEST PO#",
            "PO Description": po_description(r["order_number"], r.get("shipping_fee")),
            "PO Vendor": vendor,
            "PO Start Ship": r.get("start_date", ""), "PO End Ship": r.get("cancel_date", ""),
            "PO Received at location": RECEIVED_AT_LOCATION,
            "Item #": item_no, "PO Line Qty": int(float(r["quantity"])),
            "PO Line Unit Cost": cost, "Item Default Cost": cost,
        })
    if missing:
        print(f"  WARNING: {len(missing)} JOOR line(s) had no matching Heartland item "
              f"(vendor/style/color/size): {missing}\n  -> the vendor's items WERE "
              f"loaded, so these specific style/color/size combos aren't in Heartland: "
              f"check the grids import ran and generated items for every size.")
    write_xlsx(out_path, PO_HEADERS, out_rows, text_cols=("Item #",))
    print(f"wrote {out_path} ({len(out_rows)} PO line(s))")

# ------------------------- cli ------------------------------------------------
def main():
    ap = argparse.ArgumentParser(description="JOOR -> Heartland converter (Romi Boutique)")
    sub = ap.add_subparsers(dest="cmd", required=True)
    g = sub.add_parser("grids"); g.add_argument("joor_file"); g.add_argument("-o", default="grids_import.xlsx")
    p = sub.add_parser("po"); p.add_argument("joor_file")
    p.add_argument("-o", default="po_import.xlsx")
    a = ap.parse_args()
    rows = read_joor(a.joor_file)
    print(f"read {len(rows)} JOOR line(s) from {a.joor_file}")
    if a.cmd == "grids":
        build_grids(rows, a.o)
    else:
        build_po(rows, a.o)

if __name__ == "__main__":
    main()
