#!/usr/bin/env python3
"""ARMODILO CI CONVERTER (single file: script + settings + logo)
Turns a FedEx commercial-invoice output PDF into the Armodilo-branded CI.
Usage: python3 armodilo_ci_converter.py FEDEX.pdf OUT.pdf|OUT.docx --aci 527617844119
(.docx output is uploaded to Google Drive and converted to a Google Doc)
To add a customer or part number, edit CONFIG below.
"""

import re, json, sys, argparse, os, io, base64
from reportlab.lib.utils import ImageReader
import pdfplumber
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.lib.colors import black, red, HexColor


CONFIG = {
  "shipper": {
    "lines": [
      "SHIPPING@ARMODILO.COM",
      "ARMODILO DISPLAY SOLUTIONS",
      "350 SHIRLEY AVE, UNIT 5",
      "KITCHENER",
      "ON N2B2E1 CA (519) 206-5000"
    ],
    "id_ein": "80882 8297 RT0001"
  },
  "broker_extra_line": "FedEx Trade Network Acct#:4869598",
  "customers": {
    "COTALITY": {
      "email": "rhenz@cotality.com",
      "importer_id": "n/a"
    }
  },
  "parts": {
    "ALUR-14": {
      "title": "ALUR Tablet Kiosk 14” (*Please see attached diagram ALUR-14-ASSY-01-REV1)",
      "components": [
        {
          "material": "Aluminum",
          "value": 216.06,
          "weight": "2.9525",
          "origin": [
            [
              "Primary Country of Smelt",
              "Canada"
            ],
            [
              "Secondary Country of Smelt",
              "Canada"
            ],
            [
              "Country of Most Recent Cast",
              "Canada"
            ]
          ],
          "reference": "CBM.pdf",
          "weight_suffix": "6.51 lbs 12.5%"
        },
        {
          "material": "Aluminum",
          "value": 161.12,
          "weight": "4.7294",
          "origin": [
            [
              "Primary Country of Smelt",
              "China"
            ],
            [
              "Secondary Country of Smelt",
              "China"
            ],
            [
              "Country of Most Recent Cast",
              "China"
            ]
          ],
          "reference": "RJB.pdf",
          "weight_suffix": "10.43 lbs 20 %"
        },
        {
          "material": "Steel",
          "value": 67.78,
          "weight": "8.1591",
          "value_label": "Total Component Value",
          "weight_label": "Total Component Weight",
          "origin": [
            [
              "Country of Melt & Pour",
              "Canada"
            ]
          ],
          "reference": "Dana.pdf, CBM.pdf",
          "weight_suffix": "17.99 lbs 35%"
        }
      ],
      "item": {
        "part": "ALUR - 14",
        "hs": "8473.30.5100",
        "unit_value": 899.99,
        "desc": "ALUR Tablet Kiosk 14"
      }
    },
    "ALUR-16": {
      "title": "ALUR Tablet Kiosk 16” (*Please see attached diagram ALUR-16-ASSY-01-REV1)",
      "components": [
        {
          "material": "Aluminum",
          "value": 216.06,
          "weight": "6.8975",
          "origin": [
            [
              "Primary Country of Smelt",
              "Canada"
            ],
            [
              "Secondary Country of Smelt",
              "Canada"
            ],
            [
              "Country of Most Recent Cast",
              "Canada"
            ]
          ],
          "reference": "CBM.pdf"
        },
        {
          "material": "Aluminum",
          "value": 161.12,
          "weight": "11.9025",
          "origin": [
            [
              "Primary Country of Smelt",
              "China"
            ],
            [
              "Secondary Country of Smelt",
              "China"
            ],
            [
              "Country of Most Recent Cast",
              "China"
            ]
          ],
          "reference": "RJB.pdf"
        },
        {
          "material": "Steel",
          "value": 67.78,
          "weight": "19.75",
          "value_label": "Total Component Value",
          "weight_label": "Total Component Weight",
          "origin": [
            [
              "Country of Melt & Pour",
              "Canada"
            ]
          ],
          "reference": "Dana.pdf, CBM.pdf"
        }
      ],
      "item": {
        "part": "ALUR - 16",
        "hs": "8473.30.5100",
        "unit_value": 999.99,
        "desc": "ALUR Tablet Kiosk 16"
      }
    },
    "CURVE": {
      "title": "CURVE Desktop Tablet Display Stand (*Please see attached diagram Armodilo_Curve.pdf)",
      "components": [
        {
          "material": "Steel",
          "value": 17.66,
          "weight": "0.50",
          "value_label": "Total Component Value",
          "weight_label": "Total Component Weight",
          "origin": [
            [
              "Country of Melt & Pour",
              "United States of America"
            ]
          ],
          "reference": "SDI.pdf",
          "name": "Faceplate Cover",
          "note": "Item “F” in attached document",
          "hs": "8473.30.5100",
          "qty": 1
        },
        {
          "material": "Aluminum",
          "value": 38.5,
          "weight": "1.140",
          "origin": [
            [
              "Primary Country of Smelt",
              "China"
            ],
            [
              "Secondary Country of Smelt",
              "China"
            ],
            [
              "Country of Most Recent Cast",
              "China"
            ]
          ],
          "reference": "Henan.pdf",
          "name": "Curve Base",
          "note": "Item “B” in attached document",
          "hs": "8473.30.5100",
          "qty": 1
        }
      ],
      "item": {
        "part": "CURVE",
        "hs": "8473.30.5100",
        "unit_value": 249.99,
        "desc": "Curve - Desktop Tablet Display Stand"
      }
    },
    "FPC": {
      "title": "FPC Faceplate Cover",
      "components": [
        {
          "material": "Steel",
          "value": 17.66,
          "weight": "0.500",
          "value_label": "Total Component Value",
          "weight_label": "Total Component Weight",
          "origin": [
            [
              "Country of Melt & Pour",
              "United States of America"
            ]
          ],
          "reference": "SDI.pdf",
          "name": "Faceplate Cover",
          "hs": "8473.30.5100",
          "qty": 1
        }
      ],
      "item": {
        "part": "FPC",
        "hs": "8473.30.5100",
        "unit_value": 100.0,
        "desc": "Faceplate Cover"
      }
    },
    "FLOOR": {
      "title": "FLOOR Freestanding Tablet Display Stand (*Please see attached diagram Armodilo_Floor.pdf)",
      "components": [
        {
          "material": "Steel",
          "value": 17.66,
          "weight": "0.50",
          "value_label": "Total Component Value",
          "weight_label": "Total Component Weight",
          "origin": [
            [
              "Country of Melt & Pour",
              "United States of America"
            ]
          ],
          "reference": "SDI.pdf",
          "name": "Faceplate Cover",
          "note": "Item “F” in attached document",
          "hs": "8473.30.5100",
          "qty": 1
        },
        {
          "material": "Aluminum",
          "value": 58.6,
          "weight": "1.022",
          "origin": [
            [
              "Primary Country of Smelt",
              "United States of America"
            ],
            [
              "Secondary Country of Smelt",
              "United States of America"
            ],
            [
              "Country of Most Recent Cast",
              "United States of America"
            ]
          ],
          "reference": "Hydro.pdf",
          "name": "Aluminum Tube",
          "note": "Item “B” in attached document",
          "hs": "8473.30.5100",
          "qty": 1
        },
        {
          "material": "Aluminum",
          "value": 48.33,
          "weight": "2.562",
          "origin": [
            [
              "Primary Country of Smelt",
              "China"
            ],
            [
              "Secondary Country of Smelt",
              "China"
            ],
            [
              "Country of Most Recent Cast",
              "China"
            ]
          ],
          "reference": "Henan.pdf",
          "name": "Base Plate",
          "note": "Item “C” in attached document",
          "hs": "8473.30.5100",
          "qty": 1
        }
      ],
      "item": {
        "part": "FLOOR",
        "hs": "8473.30.5100",
        "unit_value": 449.99,
        "desc": "Freestanding Tablet Display Stand"
      }
    },
    "KBT-02": {
      "title": "KBT-02 Keyboard Tray",
      "components": [
        {
          "material": "Aluminum",
          "value": 24.39,
          "weight": "0.408",
          "origin": [
            [
              "Primary Country of Smelt",
              "China"
            ],
            [
              "Secondary Country of Smelt",
              "China"
            ],
            [
              "Country of Most Recent Cast",
              "China"
            ]
          ],
          "reference": "Henan.pdf",
          "name": "Keyboard Tray",
          "hs": "8473.30.5100",
          "qty": 1
        }
      ],
      "item": {
        "part": "KBT-02",
        "hs": "8473.30.5100",
        "unit_value": 49.99,
        "desc": "Keyboard Tray"
      }
    },
    "KEYO": {
      "title": "KEYO Wall Mount Tablet Display Stand (*Please see attached diagram Armodilo_Keyo.pdf)",
      "components": [
        {
          "material": "Steel",
          "value": 42.0,
          "weight": ".705",
          "value_label": "Total Component Value",
          "weight_label": "Total Component Weight",
          "origin": [
            [
              "Country of Melt & Pour",
              "United States of America"
            ]
          ],
          "reference": "SDI.pdf",
          "name": "Faceplate Cover",
          "note": "Item “F” in attached document",
          "hs": "8473.30.5100",
          "qty": 1
        },
        {
          "material": "Aluminum",
          "value": 35.0,
          "weight": ".795",
          "origin": [
            [
              "Primary Country of Smelt",
              "China"
            ],
            [
              "Secondary Country of Smelt",
              "China"
            ],
            [
              "Country of Most Recent Cast",
              "China"
            ]
          ],
          "reference": "BTS.pdf (page 3)",
          "name": "Box",
          "note": "Item “G” in attached document",
          "hs": "8473.30.5100",
          "qty": 1
        }
      ],
      "item": {
        "part": "KEYO",
        "hs": "8473.30.5100",
        "unit_value": 224.99,
        "desc": "Wall Mounted Tablet Display Stand"
      }
    },
    "REVOLVE": {
      "title": "REVOLVE Revolve Pivot Add-on",
      "components": [
        {
          "material": "Steel",
          "value": 8.42,
          "weight": "0.04",
          "value_label": "Total Component Value",
          "weight_label": "Total Component Weight",
          "origin": [
            [
              "Country of Melt & Pour",
              "United States of America"
            ]
          ],
          "reference": "SDI.pdf",
          "name": "Revolve Top",
          "hs": "8473.30.5100",
          "qty": 1
        },
        {
          "material": "Aluminum",
          "value": 20.41,
          "weight": "0.14",
          "origin": [
            [
              "Primary Country of Smelt",
              "Canada"
            ],
            [
              "Secondary Country of Smelt",
              "Canada"
            ],
            [
              "Country of Most Recent Cast",
              "Canada"
            ]
          ],
          "reference": "CBM.pdf",
          "name": "Revolve Base",
          "hs": "8473.30.5100",
          "qty": 1
        }
      ],
      "item": {
        "part": "REVOLVE",
        "hs": "8473.30.5100",
        "unit_value": 99.0,
        "desc": "Revolve Pivot Add-on"
      }
    },
    "SPHERE": {
      "title": "SPHERE Desktop Tablet Display Stand (*Please see attached diagram Armodilo_Sphere.pdf)",
      "components": [
        {
          "material": "Steel",
          "value": 17.66,
          "weight": "0.50",
          "value_label": "Total Component Value",
          "weight_label": "Total Component Weight",
          "origin": [
            [
              "Country of Melt & Pour",
              "United States of America"
            ]
          ],
          "reference": "SDI.pdf",
          "name": "Faceplate Cover",
          "note": "Item “F” in attached document",
          "hs": "8473.30.5100",
          "qty": 1
        },
        {
          "material": "Steel",
          "value": 18.11,
          "weight": "2.859",
          "origin": [
            [
              "Country of Melt & Pour",
              "Canada"
            ]
          ],
          "reference": "STELCO.pdf",
          "name": "Steel Base",
          "note": "Bottom of Item “B” in attached document",
          "hs": "8473.30.5100",
          "qty": 1
        }
      ],
      "item": {
        "part": "SPHERE",
        "hs": "8473.30.5100",
        "unit_value": 256.75,
        "desc": "Desktop Tablet Display Stand"
      }
    },
    "SURFACE": {
      "title": "SURFACE Mount Tablet Display Stand (*Please see attached diagram Armodilo_Surface.pdf)",
      "components": [
        {
          "material": "Steel",
          "value": 17.66,
          "weight": "0.50",
          "value_label": "Total Component Value",
          "weight_label": "Total Component Weight",
          "origin": [
            [
              "Country of Melt & Pour",
              "United States of America"
            ]
          ],
          "reference": "SDI.pdf",
          "name": "Faceplate Cover",
          "note": "Item “F” in attached document",
          "hs": "8473.30.5100",
          "qty": 1
        },
        {
          "material": "Aluminum",
          "value": 30.62,
          "weight": ".180",
          "origin": [
            [
              "Primary Country of Smelt",
              "Canada"
            ],
            [
              "Secondary Country of Smelt",
              "Canada"
            ],
            [
              "Country of Most Recent Cast",
              "Canada"
            ]
          ],
          "reference": "CanArt.pdf",
          "name": "Surface Tube",
          "note": "Item “B” in attached document",
          "hs": "8473.30.5100",
          "qty": 1
        }
      ],
      "item": {
        "part": "SURFACE",
        "hs": "8473.30.5100",
        "unit_value": 294.99,
        "desc": "Surface Mount Tablet Display Stand"
      }
    },
    "TILT": {
      "title": "TILT Wall Mount Tablet Display Stand (*Please see attached diagram Armodilo_Tilt.pdf)",
      "components": [
        {
          "material": "Steel",
          "value": 17.66,
          "weight": "0.50",
          "value_label": "Total Component Value",
          "weight_label": "Total Component Weight",
          "origin": [
            [
              "Country of Melt & Pour",
              "United States of America"
            ]
          ],
          "reference": "SDI.pdf",
          "name": "Faceplate Cover",
          "note": "Item “D” in attached document",
          "hs": "8473.30.5100",
          "qty": 1
        },
        {
          "material": "Aluminum",
          "value": 13.88,
          "weight": ".130",
          "origin": [
            [
              "Primary Country of Smelt",
              "Canada"
            ],
            [
              "Secondary Country of Smelt",
              "Canada"
            ],
            [
              "Country of Most Recent Cast",
              "Canada"
            ]
          ],
          "reference": "Almag CA.pdf",
          "name": "Tilt Tube",
          "note": "Between Items “E” and Item “F” in attached document",
          "hs": "8473.30.5100",
          "qty": 1
        }
      ],
      "item": {
        "part": "TILT",
        "hs": "8473.30.5100",
        "unit_value": 224.99,
        "desc": "Wall Mount Tablet Display Stand"
      }
    },
    "WALL": {
      "title": "WALL Wall Mount Tablet Display Stand (*Please see attached diagram Armodilo_Wall.pdf)",
      "components": [
        {
          "material": "Steel",
          "value": 17.66,
          "weight": "0.50",
          "value_label": "Total Component Value",
          "weight_label": "Total Component Weight",
          "origin": [
            [
              "Country of Melt & Pour",
              "United States of America"
            ]
          ],
          "reference": "SDI.pdf",
          "name": "Faceplate Cover",
          "note": "Item “D” in attached document",
          "hs": "8473.30.5100",
          "qty": 1
        },
        {
          "material": "Aluminum",
          "value": 14.08,
          "weight": "0.13",
          "origin": [
            [
              "Primary Country of Smelt",
              "Canada"
            ],
            [
              "Secondary Country of Smelt",
              "Canada"
            ],
            [
              "Country of Most Recent Cast",
              "Canada"
            ]
          ],
          "reference": "Almag.pdf",
          "name": "Wall Tube",
          "note": "Item “E” in attached document",
          "hs": "8473.30.5100",
          "qty": 1
        }
      ],
      "item": {
        "part": "WALL",
        "hs": "8473.30.5100",
        "unit_value": 249.0,
        "desc": "Wall Mount Tablet Display Stand"
      }
    },
    "VESA ADAPTER": {
      "title": "VESA Adapter Vesa Adapter for 75 and 100",
      "components": [
        {
          "material": "Steel",
          "value": 9.35,
          "weight": ".0618",
          "value_label": "Total Component Value",
          "weight_label": "Total Component Weight",
          "origin": [
            [
              "Country of Melt & Pour",
              "United States of America"
            ]
          ],
          "reference": "SDI.pdf",
          "hs": "8473.30.5100",
          "qty": 1
        }
      ],
      "item": {
        "part": "VESA ADAPTER",
        "hs": "8473.30.5100",
        "unit_value": 29.0,
        "desc": "Vesa Adapter for 75 and 100"
      }
    },
    "VESA": {
      "title": "VESA Wall Mount Tablet Display Stand (*Please see attached diagram Armodilo_Vesa.pdf)",
      "components": [
        {
          "material": "Steel",
          "value": 17.66,
          "weight": "0.50",
          "value_label": "Total Component Value",
          "weight_label": "Total Component Weight",
          "origin": [
            [
              "Country of Melt & Pour",
              "United States of America"
            ]
          ],
          "reference": "SDI.pdf",
          "name": "Faceplate Cover",
          "note": "Item “D” in attached document",
          "hs": "8473.30.5100",
          "qty": 1
        }
      ],
      "item": {
        "part": "VESA",
        "hs": "8473.30.5100",
        "unit_value": 100.0,
        "desc": "Wall Mounted Tablet Display Stand"
      }
    },
    "TABLET FIT KIT": {
      "title": "Tablet Fit Kit - Plastic",
      "title_color": "black",
      "components": [],
      "item": {
        "part": "TABLET FIT KIT",
        "hs": "8473.30.5100",
        "unit_value": 54.99,
        "desc": "Tablet Fit Kit - PLASTIC"
      }
    }
  }
}

SIG_B64 = "iVBORw0KGgoAAAANSUhEUgAAAF0AAAA0CAAAAADN4WuNAAADyElEQVR42rWXTWhjVRTHf29eMVaN8hBS7UiVbAIGKUXGmclAlWwy0JEuBqF20YUj3RTpIou6yMAsSjZBu+iiIllkUbCbmUFmEaWzKBg0lc6iYqhDRaJWaTXD0IGgkcn8XbwkLx8v/SB5d5N37znvx73n/M/NeYZwGwbu66ccZ9zh837jXj/wchl+pClXyymHK4NFSeAV/VBSnps90w237Bn7Q4Ax/Effs/orwBDA6p89J7Vj70anz+DT1guBs8GXrp2aPtA6/fg7CP32D5HHnx5cftb3/cNzX3wYfuXuQwgHh66/dudLAs9dyL34xsXx10+pyDD4lm1FLtrJRdLNmscqTEjs27P99KRF0AolDk+sGdK1h2TGnk9KyjU8ilREufX1Jd+Rwm223alP0jnbNiwpgnO2opuAj8I3a+aKVU/G5/bvA2BvtJZtY0Cvdgb2M+dGev/SWcM4wU2QvyFpT6xIGo1LShKouBVfxnl/Jl6QJGaPrdXqhBOoKDcYzru4J0cg53BuuwHpko2yHdAVq/V0lCRJKyEzmmtaLjtOiZUT0KUyQEoLc82GcHZtPmTOlV38M24qdKdHgsB29nCvYGYaq8V5P9aqi/vgXDeNu9OJQGFtbSEGwHRh0ZFBAMBnjqZ2Hffd5neLx9Gn8yVqmQyvS6tAVJJkluouW5kUEK+0Q/aP3/u2n2zdISVVpDwsSTJn2moVZvZbIMXj6YnaWStN18O2yYbGZjp800DCmR4eSx9vkmHjENIW4Umfi3ulWbWjR9PLi7DjiJZY1bGZ7rWdI9R5FbrS10BSUpI2FiyAWL2CJU02arIlkeEpq/b8gVutlvKFqiQtQ1K5jH1zzaYOpWCL+wZWGzob5C1JU+G4JBXYbqWvA/gi0ZqcrTFCS1UneRGr2327m5iLjQTjjWsM1uvmBRirSgIwR6JXx2vwknYkKT9bj0YTbjIIYIIZAAazu+0hCtQrblwaIyYZunXVXknOP1NrIB89z62hSzwx/Y/g7c1/4cmP336T/8Uu1ZffuVz8+e/y7/eqAETfPTdwHuCru5/Q3H0eXLv4UUdWV9lSSXsJSdO2mx8CE8s77irYzaZTIwGAmF23Xf/5atJjVko3+jAf7f7l3jo9bmuqMSnRqteN3vvIssbrhd1mKvbepaZQP/rrLl0qRui+h98eW/djP3hF/483+fovr+hPQYGf+kE3usS3Px99Z/BydKPn8TIy3Q39icwVT/fej7R23XvU07hvXvBy7+fZ9HDv/fjUPoreh7R2V2TAU72/56Vm4GCoZ/r/BJNTyO8KgqQAAAAASUVORK5CYII="
SIG_W, SIG_H = 54, 30   # signature size in points

LOGO_B64 = "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"

def _norm(s): return re.sub(r"[^A-Z0-9]", "", str(s or "").upper())
def _family(s): m = re.match(r"^[A-Z]+", str(s or "").upper().strip()); return m.group(0) if m else ""  # first word: "ALUR TABLET KIOSK 14" -> "ALUR"
def _size(s): m = re.match(r"^[A-Z]+(\d+)$", _norm(s)); return str(int(m.group(1))) if m else None
def _sizes_in(t): return [str(int(x)) for x in re.findall(r'(?:^|[^\d.])(\d{2})(?=\s*(?:["\u201d\u2033]|in\b|inch|-?inch|\b))', str(t or ""), re.I)]
def match_part(cfg, part, desc=""):
    """Returns (key, how) or (None, reason). Exact part number first, then same product family picked by size."""
    keys = list(cfg["parts"])
    for k in keys:
        if _norm(k) == _norm(part): return k, "part number"
    fam = _family(part)
    same = [k for k in keys if fam and _family(k) == fam]
    if len(same) == 1 and not _size(same[0]):
        np, nk = _norm(part), _norm(same[0])
        if (np.startswith(nk) and re.fullmatch(r"\d*", np[len(nk):])) or (nk.startswith(np) and re.fullmatch(r"\d+", nk[len(np):])):
            return same[0], "part number"
        return None, f"not sure this is {same[0]}"
    if same:
        sized = [k for k in same if _size(k)]
        wanted = {x for x in [_size(part), *_sizes_in(part), *_sizes_in(desc)] if x}
        hits = [k for k in sized if _size(k) in wanted]
        if len(hits) == 1: return hits[0], f"size {_size(hits[0])} from the invoice"
        return None, f"invoice doesn't say which {fam} it is (saved: {', '.join(same)})"
    return None, "no saved metal content for this part number"
def find_part(cfg, part, desc=""):
    k, _ = match_part(cfg, part, desc)
    return cfg["parts"][k] if k else None
VALUE_X = 198  # where page-2 answers line up (pt from page edge)
def component_fields(c, q):
    """(label, value, note, bold) for one metal component, scaled by quantity. Bold like the original."""
    f = []
    if c.get("name"): f.append(("Component", c["name"], f"({c['note']})" if c.get("note") else "", False))
    if c.get("hs"): f.append(("H.S. Code", c["hs"], "", False))
    if c.get("qty") is not None: f.append(("Quantity", f"{c['qty'] * q:g}", "", False))
    f.append(("Material", c["material"], "", True))
    f.append((c.get("value_label", "Total Value"), f"US${c['value'] * q:,.2f}", "", True))
    dec = len(str(c["weight"]).split(".")[1]) if "." in str(c["weight"]) else 0
    w = c["weight"] if q == 1 else f"{float(c['weight']) * q:.{dec}f}"
    sfx = c.get("weight_suffix", "")
    lb = re.match(r"^([\d.]+)\s*lbs\s*(.*)$", sfx)
    if lb and q != 1: sfx = f"{float(w) * 2.20462:.2f} lbs" + (f" {lb.group(2)}" if lb.group(2) else "")  # % share unchanged
    f.append((c.get("weight_label", "Total Weight"), f"{w} kg" + (f" {sfx}" if sfx else ""), "", True))
    f += [(a, b, "", True) for a, b in c["origin"]]
    f.append(("Reference", c["reference"], "", False))
    return f
def split_paren(t):
    """Split 'text (note)' so the bracketed note can be printed in blue, like the original."""
    m = re.match(r"^(.*?)( \(.*\))$", t)
    return (m.group(1), m.group(2)) if m else (t, "")
def component_rows(c, q):
    rows = []
    if c.get("name"): rows.append(f"Component: {c['name']}" + (f" ({c['note']})" if c.get("note") else ""))
    if c.get("hs"): rows.append(f"H.S. Code: {c['hs']}")
    if c.get("qty") is not None: rows.append(f"Quantity: {c['qty'] * q:g}")
    rows.append(f"Material: {c['material']}")
    rows.append(f"{c.get('value_label', 'Total Value')}: US${c['value'] * q:,.2f}")
    dec = len(str(c["weight"]).split(".")[1]) if "." in str(c["weight"]) else 0
    w = c["weight"] if q == 1 else f"{float(c['weight']) * q:.{dec}f}"
    sfx = " " + c["weight_suffix"] if q == 1 and c.get("weight_suffix") else ""
    rows.append(f"{c.get('weight_label', 'Total Weight')}: {w} kg{sfx}")
    rows += [f"{a}: {b}" for a, b in c["origin"]] + [f"Reference: {c['reference']}"]
    return rows

HERE = os.path.dirname(os.path.abspath(__file__))
FD = "/usr/share/fonts/truetype/liberation/"
pdfmetrics.registerFont(TTFont("Arial", FD + "LiberationSans-Regular.ttf"))
pdfmetrics.registerFont(TTFont("Arial-Bold", FD + "LiberationSans-Bold.ttf"))
W, H = letter
L = 14.4          # left margin used by the template
LH = 13.5         # line height

# ---------------------------------------------------------------- parsing
def lines_of(pdf_path):
    out = []
    with pdfplumber.open(pdf_path) as pdf:
        off = 0
        for p in pdf.pages:
            rows = {}
            for w in p.extract_words(x_tolerance=1.5):
                rows.setdefault(round(w["top"]), []).append(w)
            for top in sorted(rows):
                ws = sorted(rows[top], key=lambda w: w["x0"])
                out.append({"top": top + off, "words": ws,
                            "text": " ".join(w["text"] for w in ws)})
            off += p.height
    return out

def split(line, x=260):
    left = " ".join(w["text"] for w in line["words"] if w["x0"] < x)
    right = " ".join(w["text"] for w in line["words"] if w["x0"] >= x)
    return left.strip(), right.strip()

def fmt_zip(z):
    return f"{z[:5]}-{z[5:]}" if re.fullmatch(r"\d{9}", z) else z

def parse_address(lines):
    """lines = right-column strings of an address block (name, company, addr..., city, 'ST ZIP CC phone', 'ID/EIN:..')"""
    lines = [l for l in lines if l]
    idein = ""
    if lines and lines[-1].startswith("ID/EIN:"):
        idein = lines.pop()[7:].strip()
    m = re.match(r"^([A-Z]{2})\s+(\S+)\s+([A-Z]{2})\s*(.*)$", lines[-1]) if lines else None
    state = zipc = country = phone = ""
    if m:
        state, zipc, country, phone = m.group(1), fmt_zip(m.group(2)), m.group(3), m.group(4).strip()
        lines.pop()
    city = lines.pop() if lines else ""
    return {"name": lines[0] if lines else "", "company": lines[1] if len(lines) > 1 else "",
            "street": lines[2:], "city": city, "state": state, "zip": zipc,
            "country": country, "phone": phone, "id": idein}

def parse(pdf_path):
    ls = lines_of(pdf_path)
    txt = "\n".join(l["text"] for l in ls)
    is_rule = lambda l: re.fullmatch(r"[-_]{5,}", l["text"].replace(" ", "")) is not None
    idx = lambda pat: next(i for i, l in enumerate(ls) if re.search(pat, l["text"]))
    d = {}

    i = idx(r"^Invoice Number")
    w = ls[i + 1]["text"].split()
    d["purpose"], d["currency"], d["ult_dest"] = w[0], w[1], w[2]
    i = idx(r"^Export Dt")
    w = ls[i + 1]["text"].split()
    d["export_date"], d["ci_ref"], d["pkgs"], d["bill_tc"] = w[0], w[1], w[2], w[3]
    nxt = ls[i + 2]["text"].strip()
    d["pkg_type"] = nxt if re.fullmatch(r"[A-Z]+", nxt) else ""

    def block(start_pat):
        s = idx(start_pat) + 1
        e = next(j for j in range(s, len(ls)) if is_rule(ls[j]))
        return ls[s:e]
    cons = block(r"^Shipper:")
    d["consignee"] = parse_address([split(l)[1] for l in cons])
    imp = block(r"^Broker:")
    d["broker"] = [split(l)[0] for l in imp if split(l)[0]]
    d["importer"] = parse_address([split(l)[1] for l in imp])

    # line items
    items = []
    starts = [i for i, l in enumerate(ls) if l["text"].startswith("Part Nbr:")]
    for s in starts:
        e = next(j for j in range(s, len(ls)) if is_rule(ls[j]))
        b = " ".join(l["text"] for l in ls[s:e])
        g = lambda p, default="": (re.search(p, b).group(1).strip() if re.search(p, b) else default)
        items.append({
            "part": g(r"Part Nbr:\s*(.*?)\s*Marks/Nbrs:") or g(r"Part Nbr:\s*(\S+)"),
            "marks": g(r"Marks/Nbrs:\s*(.*?)\s*C/T"),
            "mfr": g(r"C/T MFR:\s*(\S+)"),
            "net_wgt": g(r"Net\s+Wgt:\s*([\d.]+)"),
            "wgt_unit": g(r"Net\s+Wgt:\s*[\d.]+\s*(\w+)", "lbs"),
            "hs": g(r"HS Code:\s*(\S+)"),
            "qty": g(r"Unit Qty:\s*([\d.]+)"),
            "uom": g(r"Unit Qty:\s*[\d.]+\s*(\w+)", "EA"),
            "commodity_value": g(r"Commodity Value:\s*([\d.,]+)"),
            "unit_value": g(r"(?<![\d.])(\d+\.\d{6})(?![\d.])"),
            "desc": g(r"Desc:\s*(.*)$"),
        })
    d["items"] = items

    g = lambda p, default="": (re.search(p, txt, re.S).group(1).strip() if re.search(p, txt, re.S) else default)
    d["total_weight"] = g(r"Total Shipment\s+([\d.]+)")
    d["total_commodity"] = g(r"Total Commodity\s*\n?([\d.,]+)")
    d["terms"] = g(r"Terms of Sale:\s*\n?(\w+)")
    d["freight"] = g(r"Freight:\s*\n?([\d.,]+)")
    d["insurance"] = g(r"Insurance:\s*\n?([\d.,]+)")
    d["others"] = g(r"Others:\s*\n?([\d.,]+)")
    d["total_invoice"] = g(r"Total Invoice Value:\s*\n?([\d.,]+)")
    d["comments"] = re.findall(r"^\d\).*$", g(r"Comments:\s*\n(.*?)Tracking Numbers", ), re.M)
    d["tracking"] = g(r"Tracking Numbers:\s*([\d ,]+)").split()
    d["sig_date"] = g(r"Signature of shipper:.*?(\d{1,2}/\d{1,2}/\d{4})")
    # "Air Waybill Nbr" (when present) is the ACI number: same line or the line below
    d["awb"] = ""
    ai = next((i for i, l in enumerate(ls) if re.search(r"air\s*way\s*bill", l["text"], re.I)), None)
    if ai is not None:
        pick = lambda t: (lambda m: re.sub(r"[\s-]", "", m.group(0)) if m else "")(re.search(r"\b\d{4}[ -]\d{4}[ -]\d{4}\b|\b\d{9,15}\b", t))
        after = re.sub(r"(?i)^.*?air\s*way\s*bill\s*(?:nbr|no|number|#)?\.?\s*:?", "", ls[ai]["text"])
        d["awb"] = pick(after) or (pick(ls[ai + 1]["text"]) if ai + 1 < len(ls) else "")
    return d

# ---------------------------------------------------------------- rendering
class Doc:
    def __init__(self, path):
        self.c = canvas.Canvas(path, pagesize=letter)
        self.y = 0
    def t(self, x, top, s, font="Arial", size=10, color=black):
        c = self.c; c.setFont(font, size); c.setFillColor(color)
        c.drawString(x, H - top - size * 0.905, s)
    def tc(self, top, s, font="Arial-Bold", size=12, color=black):
        c = self.c; c.setFont(font, size); c.setFillColor(color)
        c.drawCentredString(W / 2, H - top - size * 0.905, s)
    def tr(self, xr, top, s, size=10):
        self.c.setFont("Arial", size); self.c.setFillColor(black)
        self.c.drawRightString(xr, H - top - size * 0.905, s)
    def rule(self, top, char="-"):
        c = self.c; c.setStrokeColor(black); c.setLineWidth(0.6)
        yy = H - top - 6
        if char == "-": c.setDash(3, 2)
        else: c.setDash()
        c.line(L, yy, W - L, yy); c.setDash()

def money(s):
    return f"{float(str(s).replace(',', '')):,.2f}"

def render(d, cfg, out, aci):
    doc = Doc(out); c = doc.c
    cust = cfg["customers"].get(d["importer"]["company"].upper(), {})
    parts = [it for it in d["items"] if find_part(cfg, it["part"], it["desc"])]
    has_p2 = bool(parts)
    total_pages = 2

    c.drawImage(ImageReader(io.BytesIO(base64.b64decode(LOGO_B64))), 225, H - 16 - 36.6, width=161, height=36.6, mask="auto")
    doc.tc(60, f"ACI-{aci}")
    doc.tc(76.5, "Commercial Invoice")
    if True:  # always shown
        doc.tc(92.8, "Please see Page 2 for Steel/Aluminum Content Breakdown", size=11, color=red)

    y = 107.6
    doc.t(L, y, "Invoice Number"); doc.t(95.5, y, "Purpose of Shipment"); doc.t(302.7, y, "Curr"); doc.t(355.5, y, "Ult Dest.")
    y += LH
    doc.t(122.4, y, d["purpose"]); doc.t(300.7, y, d["currency"]); doc.t(355.1, y, d["ult_dest"])
    y += LH
    doc.t(L, y, "Export Dt"); doc.t(81.1, y, "C.I. References"); doc.t(275.5, y, "Pkgs"); doc.t(333.8, y, "Bill T/C")
    y += LH
    doc.t(L, y, d["export_date"]); doc.t(86.6, y, d["ci_ref"]); doc.t(222.8, y, d["pkg_type"])
    doc.t(291.1, y, d["pkgs"]); doc.t(332.8, y, d["bill_tc"])
    y += LH; doc.rule(y); y += LH

    # shipper / consignee
    cn = d["consignee"]
    doc.t(L, y, "Shipper:"); doc.t(302.4, y, "Consignee:")
    right = [cn["name"], cn["company"], *cn["street"],
             f"{cn['city']},  {cn['state']} {cn['zip']} {cn['country']}".strip(), cn["phone"]]
    left = cfg["shipper"]["lines"]
    for k in range(max(len(left), len(right))):
        y += LH
        if k < len(left): doc.t(L, y, left[k])
        if k < len(right): doc.t(302.4, y, right[k])
    y += LH
    doc.t(L, y, f"ID/EIN: {cfg['shipper']['id_ein']}"); doc.t(268.1, y, "ID/EIN:")
    y += LH; doc.rule(y); y += LH

    # broker / importer
    im = d["importer"]
    doc.t(L, y, "Broker:"); doc.t(266.4, y, "Importer:")
    bl = d["broker"] + [cfg["broker_extra_line"]]
    ir = [im["name"], im["company"], *im["street"], f"{im['city']},  {im['state']} {im['zip']} {im['country']}".strip()]
    for k in range(max(len(bl), len(ir))):
        y += LH
        if k < len(bl): doc.t(L, y, bl[k])
        if k < len(ir): doc.t(266.4, y, ir[k])
    y += LH * 2
    imp_id = im["id"] or cn["id"] or cust.get("importer_id", "n/a")
    email = cust.get("email", "")
    doc.tc(y, "  ".join(s for s in [f"ID/EIN: {imp_id}", email, im["phone"] or cn["phone"]] if s), size=11, font="Arial-Bold")
    y += 15; doc.rule(y); y += LH

    for it in d["items"]:
        doc.t(L, y, f"Part Nbr:{it['part']}"); doc.t(230.5, y, f"Marks/Nbrs:{(' ' + it['marks']) if it['marks'] else ''}")
        doc.t(328.8, y, f"C/T MFR:{it['mfr']}"); doc.t(413.7, y, "Net Wgt:"); doc.t(474.8, y, it["net_wgt"]); doc.t(511.0, y, it["wgt_unit"])
        y += LH
        doc.t(L, y, f"HS Code: {it['hs']}"); doc.t(185.5, y, f"Unit Qty: {it['qty']}"); doc.t(252.2, y, it["uom"])
        doc.t(307.2, y, "Unit Value:"); doc.t(402.6, y, f"Commodity Value: {it['commodity_value']}")
        y += LH; doc.t(302.4, y, it["unit_value"])
        y += LH; doc.t(L, y, f"Desc: {it['desc']}")
        y += LH; doc.rule(y); y += LH

    doc.t(194.9, y, "Total Shipment"); doc.t(297.1, y, d["total_weight"]); doc.t(363.8, y, "Total Commodity"); doc.t(468.8, y, d["total_commodity"])
    y += LH; doc.t(230.4, y, "Weight:"); doc.t(308.5, y, "lbs"); doc.t(410.2, y, "Value:")
    rx = 528
    for lab, val in [("Terms of Sale:", d["terms"]), ("Freight:", d["freight"]), ("Insurance:", d["insurance"]), ("Others:", d["others"])]:
        y += LH; doc.tr(410.0, y, lab); doc.t(482.4, y, val)
    y += LH; doc.tr(410.0, y, "Total Invoice Value:"); doc.t(452.0, y, f"US${d['total_invoice']}" if d["currency"] == "USD" else d["total_invoice"])
    y += LH * 2
    doc.t(L, y, "Comments:")
    for cm in d["comments"]:
        y += LH; doc.t(L, y, cm)
    y += LH; doc.t(L, y, "Tracking Numbers:"); doc.t(108, y, ", ".join(d["tracking"]))
    y += LH * 2; doc.t(L, y, "I declare all information in this invoice to be true and correct.")
    y += LH + 17   # room for the signature, which rises above the text line
    lab = "Signature of shipper: "; lw = pdfmetrics.stringWidth(lab, "Arial", 10)
    doc.t(L, y, lab)
    c.drawImage(ImageReader(io.BytesIO(base64.b64decode(SIG_B64))), L + lw, H - y - 9 - SIG_H + 6, width=SIG_W, height=SIG_H)
    doc.t(L + lw + SIG_W + 3, y, "SHIPPING@ARMODILO.COM"); doc.t(348, y, d["sig_date"])
    y += LH * 2; doc.rule(y, "_")
    y += LH * 2; doc.t(L, y, f"Page 1 of {total_pages}")

    if has_p2:
        c.showPage()
        doc.tc(16.5, "Metal Content", color=red)
        y = 49.1
        for n, it in enumerate(d["items"], 1):
            p = find_part(cfg, it["part"], it["desc"])
            if not p: continue
            q = float(it["qty"] or 1)
            t1, t2 = split_paren(f"Line {n}: {p['title']}")
            doc.t(L, y, t1, font="Arial-Bold", color=black if p.get("title_color") == "black" else red)
            if t2: doc.t(L + pdfmetrics.stringWidth(t1 + "  ", "Arial-Bold", 10), y, t2.strip(), font="Arial-Bold", color=HexColor("#0000FF"))
            y += LH * 2
            for comp in p["components"]:
                for lab, val, note, bold in component_fields(comp, q):
                    if y > H - 40:
                        c.showPage(); y = 30
                    fnt = "Arial-Bold" if bold else "Arial"
                    doc.t(50.4, y, lab + ":", color=red)
                    doc.t(VALUE_X, y, val, font=fnt, color=red)
                    if note: doc.t(VALUE_X + pdfmetrics.stringWidth(val + "  ", fnt, 10), y, note, color=HexColor("#0000FF"))
                    y += LH
                y += LH
        doc.t(L, y, "Page 2 of 2")
    c.save()

# ---------------------------------------------------------------- Google Doc / Word output

def _slim_docx(path):
    """Shrink the python-docx output (huge default styles) so it uploads quickly to Google Drive."""
    import zipfile, re as _re
    from PIL import Image
    zin = zipfile.ZipFile(path); items = {n: zin.read(n) for n in zin.namelist()}; zin.close()
    items.pop("word/stylesWithEffects.xml", None)
    for n in list(items):
        if n.startswith(("customXml/", "docProps/thumbnail", "word/numbering", "word/theme", "word/webSettings", "word/fontTable")):
            items.pop(n)
    items["word/_rels/document.xml.rels"] = _re.sub(rb'<Relationship [^>]*(customXml|numbering|theme|webSettings|fontTable)[^>]*/>', b"", items["word/_rels/document.xml.rels"])
    items["_rels/.rels"] = _re.sub(rb'<Relationship [^>]*thumbnail[^>]*/>', b"", items["_rels/.rels"])
    items["[Content_Types].xml"] = _re.sub(rb'<Override [^>]*(customXml|numbering|theme|webSettings|fontTable|thumbnail)[^>]*/>', b"", items["[Content_Types].xml"])
    items["word/_rels/document.xml.rels"] = _re.sub(rb'<Relationship [^>]*stylesWithEffects[^>]*/>', b"", items["word/_rels/document.xml.rels"])
    items["[Content_Types].xml"] = _re.sub(rb'<Override [^>]*stylesWithEffects[^>]*/>', b"", items["[Content_Types].xml"])
    items["word/styles.xml"] = (b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        b'<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
        b'<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:eastAsia="Arial" w:cs="Arial"/>'
        b'<w:sz w:val="20"/><w:szCs w:val="20"/></w:rPr></w:rPrDefault><w:pPrDefault><w:pPr>'
        b'<w:spacing w:before="0" w:after="0" w:line="270" w:lineRule="exact"/></w:pPr></w:pPrDefault></w:docDefaults>'
        b'<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:qFormat/></w:style>'
        b'</w:styles>')
    for n in list(items):
        if n.startswith("word/media/") and n.endswith(".png"):
            b = io.BytesIO(); im_ = Image.open(io.BytesIO(items[n]))
            (im_ if im_.mode == "L" else im_.convert("RGB").quantize(16)).save(b, "PNG", optimize=True)
            items[n] = b.getvalue()
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        for n, data in items.items(): z.writestr(n, data)

def render_docx(d, cfg, out, aci):
    from docx import Document
    from docx.shared import Pt, RGBColor, Inches
    from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_TAB_ALIGNMENT
    from docx.oxml.ns import qn
    from docx.oxml import OxmlElement
    doc = Document()
    sec = doc.sections[0]
    sec.page_width, sec.page_height = Inches(8.5), Inches(11)
    sec.left_margin = sec.right_margin = Pt(14.4)
    sec.top_margin, sec.bottom_margin = Pt(16), Pt(30)
    st = doc.styles["Normal"]; st.font.name = "Arial"; st.font.size = Pt(10)
    st.element.rPr.rFonts.set(qn("w:eastAsia"), "Arial")
    st.paragraph_format.space_after = Pt(0); st.paragraph_format.space_before = Pt(0)
    st.paragraph_format.line_spacing = Pt(13.5)
    RED = RGBColor(0xFF, 0, 0)

    def P(segs=(), bold=False, color=None, size=None, center=False, left=0):
        """segs: list of (x_pt_on_page, text) or (x, text, 'R') for right-aligned tab ending at x"""
        p = doc.add_paragraph()
        if center: p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        if left: p.paragraph_format.left_indent = Pt(left)
        for i, s in enumerate(segs):
            x, text = s[0], s[1]
            if i or (x > L + 1 and not center and not left):
                p.paragraph_format.tab_stops.add_tab_stop(
                    Pt(x - L), WD_TAB_ALIGNMENT.RIGHT if len(s) > 2 else WD_TAB_ALIGNMENT.LEFT)
                r = p.add_run("\t")
            r = p.add_run(text); r.bold = bold
            if color: r.font.color.rgb = color
            if size: r.font.size = Pt(size)
        return p
    def rule(style="dashed"):
        p = doc.add_paragraph(); p.paragraph_format.line_spacing = Pt(6)
        pPr = p._p.get_or_add_pPr(); b = OxmlElement("w:pBdr"); bt = OxmlElement("w:bottom")
        bt.set(qn("w:val"), style); bt.set(qn("w:sz"), "6"); bt.set(qn("w:space"), "1"); bt.set(qn("w:color"), "000000")
        b.append(bt); pPr.append(b)
        doc.add_paragraph().paragraph_format.line_spacing = Pt(6)

    cust = cfg["customers"].get(d["importer"]["company"].upper(), {})
    cn, im = d["consignee"], d["importer"]
    has_p2 = any(find_part(cfg, it["part"], it["desc"]) for it in d["items"])

    lp = doc.add_paragraph(); lp.alignment = WD_ALIGN_PARAGRAPH.CENTER
    lp.paragraph_format.line_spacing = 1.0
    lp.add_run().add_picture(io.BytesIO(base64.b64decode(LOGO_B64)), width=Pt(161))
    P([(0, f"ACI-{aci}")], bold=True, size=12, center=True)
    P([(0, "Commercial Invoice")], bold=True, size=12, center=True)
    P([(0, "Please see Page 2 for Steel/Aluminum Content Breakdown")], bold=True, size=11, color=RED, center=True)
    P([(L, "Invoice Number"), (95.5, "Purpose of Shipment"), (302.7, "Curr"), (355.5, "Ult Dest.")])
    P([(122.4, d["purpose"]), (300.7, d["currency"]), (355.1, d["ult_dest"])])
    P([(L, "Export Dt"), (81.1, "C.I. References"), (275.5, "Pkgs"), (333.8, "Bill T/C")])
    P([(L, d["export_date"]), (86.6, d["ci_ref"]), (222.8, d["pkg_type"]), (291.1, d["pkgs"]), (332.8, d["bill_tc"])])
    rule()
    left = ["Shipper:"] + cfg["shipper"]["lines"]
    right = ["Consignee:", cn["name"], cn["company"], *cn["street"],
             f"{cn['city']},  {cn['state']} {cn['zip']} {cn['country']}", cn["phone"]]
    for k in range(max(len(left), len(right))):
        segs = []
        if k < len(left): segs.append((L, left[k]))
        if k < len(right): segs.append((302.4, right[k]))
        P(segs)
    P([(L, f"ID/EIN: {cfg['shipper']['id_ein']}"), (268.1, "ID/EIN:")])
    rule()
    bl = ["Broker:"] + d["broker"] + [cfg["broker_extra_line"]]
    ir = ["Importer:", im["name"], im["company"], *im["street"], f"{im['city']},  {im['state']} {im['zip']} {im['country']}"]
    for k in range(max(len(bl), len(ir))):
        segs = []
        if k < len(bl): segs.append((L, bl[k]))
        if k < len(ir): segs.append((266.4, ir[k]))
        P(segs)
    P()
    imp_id = im["id"] or cn["id"] or cust.get("importer_id", "n/a")
    P([(0, "  ".join(s for s in [f"ID/EIN: {imp_id}", cust.get("email", ""), im["phone"] or cn["phone"]] if s))], bold=True, size=11, center=True)
    rule()
    for it in d["items"]:
        P([(L, f"Part Nbr:{it['part']}"), (230.5, "Marks/Nbrs:" + (" " + it["marks"] if it["marks"] else "")), (328.8, f"C/T MFR:{it['mfr']}"),
           (413.7, "Net Wgt:"), (474.8, it["net_wgt"]), (511.0, it["wgt_unit"])])
        P([(L, f"HS Code: {it['hs']}"), (185.5, f"Unit Qty: {it['qty']}"), (252.2, it["uom"]), (307.2, "Unit Value:"), (402.6, f"Commodity Value: {it['commodity_value']}")])
        P([(302.4, it["unit_value"])])
        P([(L, f"Desc: {it['desc']}")])
        rule()
    P([(194.9, "Total Shipment"), (297.1, d["total_weight"]), (363.8, "Total Commodity"), (468.8, d["total_commodity"])])
    P([(230.4, "Weight:"), (308.5, "lbs"), (410.2, "Value:")])
    for lab, val in [("Terms of Sale:", d["terms"]), ("Freight:", d["freight"]), ("Insurance:", d["insurance"]), ("Others:", d["others"])]:
        P([(410.0, lab, "R"), (482.4, val)])
    P([(410.0, "Total Invoice Value:", "R"), (452.0, f"US${d['total_invoice']}" if d["currency"] == "USD" else d["total_invoice"])])
    P()
    P([(L, "Comments:")])
    for cm in d["comments"]: P([(L, cm)])
    P([(L, "Tracking Numbers:"), (108, ", ".join(d["tracking"]))])
    P()
    P([(L, "I declare all information in this invoice to be true and correct.")])
    P([(L, "Signature of shipper: SHIPPING@ARMODILO.COM"), (348, d["sig_date"])])
    rule("single")
    P([(L, "Page 1 of 2")])

    from docx.enum.text import WD_BREAK
    doc.paragraphs[-1].add_run().add_break(WD_BREAK.PAGE)
    P([(0, "Metal Content")], bold=True, size=12, color=RED, center=True)
    P()
    for n, it in enumerate(d["items"], 1):
        p = find_part(cfg, it["part"], it["desc"])
        P([(L, f"Line {n}: {it['part']} {it['desc']} (*Please see attached diagram {p['diagram'] if p else '[diagram]'})")], bold=True, color=RED)
        P()
        if not p:
            P([(0, "[Metal content for this part still to be added]")], color=RED, left=36)
            P(); continue
        q = float(it["qty"] or 1)
        for comp in p["components"]:
            rows = component_rows(comp, q)
            for r in rows: P([(0, r)], color=RED, left=36)
            P()
    P([(L, "Page 2 of 2")])
    doc.save(out)
    _slim_docx(out)


# ---------------------------------------------------------------- Word / Google Doc output

def _slim_docx(path):
    """Shrink the python-docx output (huge default styles) so it uploads quickly to Google Drive."""
    import zipfile, re as _re
    from PIL import Image
    zin = zipfile.ZipFile(path); items = {n: zin.read(n) for n in zin.namelist()}; zin.close()
    items.pop("word/stylesWithEffects.xml", None)
    for n in list(items):
        if n.startswith(("customXml/", "docProps/thumbnail", "word/numbering", "word/theme", "word/webSettings", "word/fontTable")):
            items.pop(n)
    items["word/_rels/document.xml.rels"] = _re.sub(rb'<Relationship [^>]*(customXml|numbering|theme|webSettings|fontTable)[^>]*/>', b"", items["word/_rels/document.xml.rels"])
    items["_rels/.rels"] = _re.sub(rb'<Relationship [^>]*thumbnail[^>]*/>', b"", items["_rels/.rels"])
    items["[Content_Types].xml"] = _re.sub(rb'<Override [^>]*(customXml|numbering|theme|webSettings|fontTable|thumbnail)[^>]*/>', b"", items["[Content_Types].xml"])
    items["word/_rels/document.xml.rels"] = _re.sub(rb'<Relationship [^>]*stylesWithEffects[^>]*/>', b"", items["word/_rels/document.xml.rels"])
    items["[Content_Types].xml"] = _re.sub(rb'<Override [^>]*stylesWithEffects[^>]*/>', b"", items["[Content_Types].xml"])
    items["word/styles.xml"] = (b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        b'<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
        b'<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:eastAsia="Arial" w:cs="Arial"/>'
        b'<w:sz w:val="20"/><w:szCs w:val="20"/></w:rPr></w:rPrDefault><w:pPrDefault><w:pPr>'
        b'<w:spacing w:before="0" w:after="0" w:line="270" w:lineRule="exact"/></w:pPr></w:pPrDefault></w:docDefaults>'
        b'<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:qFormat/></w:style>'
        b'</w:styles>')
    for n in list(items):
        if n.startswith("word/media/") and n.endswith(".png"):
            b = io.BytesIO(); im_ = Image.open(io.BytesIO(items[n]))
            (im_ if im_.mode == "L" else im_.convert("RGB").quantize(16)).save(b, "PNG", optimize=True)
            items[n] = b.getvalue()
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        for n, data in items.items(): z.writestr(n, data)

def render_docx(d, cfg, out, aci):
    from docx import Document
    from docx.shared import Pt, RGBColor, Inches
    from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_TAB_ALIGNMENT, WD_LINE_SPACING
    from docx.oxml.ns import qn
    from docx.oxml import OxmlElement
    RED = RGBColor(0xFF, 0, 0)
    doc = Document()
    sec = doc.sections[0]
    sec.page_width, sec.page_height = Inches(8.5), Inches(11)
    sec.left_margin = sec.right_margin = Pt(L); sec.top_margin = Pt(16); sec.bottom_margin = Pt(20)
    st = doc.styles["Normal"]; st.font.name = "Arial"; st.font.size = Pt(10)
    st.element.rPr.rFonts.set(qn("w:eastAsia"), "Arial")
    pf = st.paragraph_format; pf.space_before = pf.space_after = Pt(0)
    pf.line_spacing_rule = WD_LINE_SPACING.EXACTLY; pf.line_spacing = Pt(LH)

    def para(segs=(), center=None, bold=False, size=10, color=None, rule=None, page_break=False, extra=None):
        """segs: list of (x_points, text) or (x, text, 'right'); center: centered text; extra: (text, color) run at the end"""
        pp = doc.add_paragraph()
        if page_break:
            from docx.enum.text import WD_BREAK
            pp.add_run().add_break(WD_BREAK.PAGE)
        if center is not None:
            pp.alignment = WD_ALIGN_PARAGRAPH.CENTER
            r = pp.add_run(center); r.bold = bold; r.font.size = Pt(size)
            if color: r.font.color.rgb = color
            if size > 10: pp.paragraph_format.line_spacing = Pt(size * 1.4)
        for k, seg in enumerate(segs):
            x, text = seg[0], seg[1]
            align = seg[2] if len(seg) > 2 else "left"
            if x > L + 0.5:
                pp.paragraph_format.tab_stops.add_tab_stop(Pt(x - L),
                    WD_TAB_ALIGNMENT.RIGHT if align == "right" else WD_TAB_ALIGNMENT.LEFT)
                pp.add_run("\t")
            r = pp.add_run(text); r.bold = bold; r.font.size = Pt(size)
            if color: r.font.color.rgb = color
        if extra and extra[0]:
            r = pp.add_run(extra[0]); r.bold = bold; r.font.size = Pt(size); r.font.color.rgb = extra[1]
        if rule:
            pPr = pp._p.get_or_add_pPr(); bdr = OxmlElement("w:pBdr"); b = OxmlElement("w:bottom")
            b.set(qn("w:val"), rule); b.set(qn("w:sz"), "6"); b.set(qn("w:space"), "1"); b.set(qn("w:color"), "000000")
            bdr.append(b); pPr.append(bdr)
        return pp
    blank = lambda: para()
    def hr(kind="dashed"):
        pp = para(rule=kind); pp.paragraph_format.line_spacing = Pt(7)
        para().paragraph_format.line_spacing = Pt(6.5)

    lp = doc.add_paragraph(); lp.alignment = WD_ALIGN_PARAGRAPH.CENTER
    lp.paragraph_format.line_spacing_rule = WD_LINE_SPACING.SINGLE
    lp.add_run().add_picture(io.BytesIO(base64.b64decode(LOGO_B64)), width=Pt(161))
    para(center=f"ACI-{aci}", bold=True, size=12)
    para(center="Commercial Invoice", bold=True, size=12)
    para(center="Please see Page 2 for Steel/Aluminum Content Breakdown", bold=True, size=11, color=RED)
    para([(L, "Invoice Number"), (95.5, "Purpose of Shipment"), (302.7, "Curr"), (355.5, "Ult Dest.")])
    para([(122.4, d["purpose"]), (300.7, d["currency"]), (355.1, d["ult_dest"])])
    para([(L, "Export Dt"), (81.1, "C.I. References"), (275.5, "Pkgs"), (333.8, "Bill T/C")])
    para([(L, d["export_date"]), (86.6, d["ci_ref"]), (222.8, d["pkg_type"]), (291.1, d["pkgs"]), (332.8, d["bill_tc"])])
    hr()
    cn, im = d["consignee"], d["importer"]
    cust = cfg["customers"].get(im["company"].upper(), {})
    para([(L, "Shipper:"), (302.4, "Consignee:")])
    left = cfg["shipper"]["lines"]
    right = [cn["name"], cn["company"], *cn["street"], f"{cn['city']},  {cn['state']} {cn['zip']} {cn['country']}", cn["phone"]]
    for k in range(max(len(left), len(right))):
        segs = []
        if k < len(left): segs.append((L, left[k]))
        if k < len(right): segs.append((302.4, right[k]))
        para(segs)
    para([(L, f"ID/EIN: {cfg['shipper']['id_ein']}"), (268.1, "ID/EIN:")])
    hr()
    para([(L, "Broker:"), (266.4, "Importer:")])
    bl = d["broker"] + [cfg["broker_extra_line"]]
    ir = [im["name"], im["company"], *im["street"], f"{im['city']},  {im['state']} {im['zip']} {im['country']}"]
    for k in range(max(len(bl), len(ir))):
        segs = []
        if k < len(bl): segs.append((L, bl[k]))
        if k < len(ir): segs.append((266.4, ir[k]))
        para(segs)
    blank()
    imp_id = im["id"] or cn["id"] or cust.get("importer_id", "n/a")
    para(center="  ".join(x for x in [f"ID/EIN: {imp_id}", cust.get("email", ""), im["phone"] or cn["phone"]] if x), bold=True, size=11)
    hr()
    for it in d["items"]:
        para([(L, f"Part Nbr:{it['part']}"), (230.5, "Marks/Nbrs:" + ((" " + it["marks"]) if it["marks"] else "")),
              (328.8, f"C/T MFR:{it['mfr']}"), (413.7, "Net Wgt:"), (474.8, it["net_wgt"]), (511.0, it["wgt_unit"])])
        para([(L, f"HS Code: {it['hs']}"), (185.5, f"Unit Qty: {it['qty']}"), (252.2, it["uom"]),
              (307.2, "Unit Value:"), (402.6, f"Commodity Value: {it['commodity_value']}")])
        para([(302.4, it["unit_value"])])
        para([(L, f"Desc: {it['desc']}")])
        hr()
    para([(194.9, "Total Shipment"), (297.1, d["total_weight"]), (363.8, "Total Commodity"), (468.8, d["total_commodity"])])
    para([(230.4, "Weight:"), (308.5, "lbs"), (410.2, "Value:")])
    for lab, val in [("Terms of Sale:", d["terms"]), ("Freight:", d["freight"]), ("Insurance:", d["insurance"]), ("Others:", d["others"])]:
        para([(410.0, lab, "right"), (482.4, val)])
    para([(410.0, "Total Invoice Value:", "right"), (452.0, f"US${d['total_invoice']}" if d["currency"] == "USD" else d["total_invoice"])])
    blank(); para([(L, "Comments:")])
    for cm in d["comments"]: para([(L, cm)])
    para([(L, "Tracking Numbers:"), (108, ", ".join(d["tracking"]))])
    blank(); para([(L, "I declare all information in this invoice to be true and correct.")])
    sp = doc.add_paragraph(); sp.paragraph_format.line_spacing_rule = WD_LINE_SPACING.SINGLE
    sp.paragraph_format.tab_stops.add_tab_stop(Pt(348 - L))
    sp.add_run("Signature of shipper: ")
    sp.add_run().add_picture(io.BytesIO(base64.b64decode(SIG_B64)), width=Pt(SIG_W), height=Pt(SIG_H))
    sp.add_run(" SHIPPING@ARMODILO.COM\t" + d["sig_date"])
    blank(); hr("single"); blank()
    para([(L, "Page 1 of 2")])

    # page 2
    para(center="Metal Content", bold=True, size=12, color=RED, page_break=True)
    blank()
    for n, it in enumerate(d["items"], 1):
        p = find_part(cfg, it["part"], it["desc"])
        if not p:
            para([(L, f"Line {n}: {it['part']} {it['desc']} -- METAL CONTENT TO BE ADDED")], bold=True, color=RED); blank(); continue
        q = float(it["qty"] or 1)
        BLUE = RGBColor(0, 0, 0xFF)
        t1, t2 = split_paren(f"Line {n}: {p['title']}")
        para([(L, t1)], bold=True, color=RGBColor(0, 0, 0) if p.get('title_color') == 'black' else RED, extra=((' ' + t2) if t2 else '', BLUE))
        blank()
        for comp in p["components"]:
            for lab, val, note, bold in component_fields(comp, q):
                pp = doc.add_paragraph()
                pp.paragraph_format.tab_stops.add_tab_stop(Pt(50.4 - L)); pp.paragraph_format.tab_stops.add_tab_stop(Pt(VALUE_X - L))
                for text, b, col in (("\t" + lab + ":", False, RED), ("\t" + val, bold, RED), (("  " + note) if note else "", False, BLUE)):
                    if text:
                        r = pp.add_run(text); r.bold = b; r.font.size = Pt(10); r.font.color.rgb = col
            blank()
    para([(L, "Page 2 of 2")])
    doc.save(out)
    _slim_docx(out)

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("fedex_pdf"); ap.add_argument("out_pdf"); ap.add_argument("--aci")
    ap.add_argument("--config", default=None)
    ap.add_argument("--docx", help="also write a Word/Google-Doc version to this path")
    a = ap.parse_args()
    cfg = json.load(open(a.config)) if a.config else CONFIG
    d = parse(a.fedex_pdf)
    aci = a.aci or d.get("awb") or (d["tracking"][0] if d["tracking"] else "XXXXXXXXXXXX")
    warn = []
    if not a.aci and not d.get("awb"): warn.append(f"ACI number defaulted to tracking number {aci} (no Air Waybill Nbr on the file)")
    if d["importer"]["company"].upper() not in cfg["customers"]:
        warn.append(f"No email on file for importer '{d['importer']['company']}'")
    for it in d["items"]:
        if not find_part(cfg, it["part"], it["desc"]):
            warn.append(f"Line {it['part']}: {match_part(cfg, it['part'], it['desc'])[1]} -- page 2 says METAL CONTENT TO BE ADDED")
    render_docx(d, cfg, a.out_pdf, aci) if a.out_pdf.lower().endswith(".docx") else render(d, cfg, a.out_pdf, aci)
    if a.docx: render_docx(d, cfg, a.docx, aci)
    print(json.dumps(d, indent=1)); print("WARNINGS:", warn)

if __name__ == "__main__":
    main()
