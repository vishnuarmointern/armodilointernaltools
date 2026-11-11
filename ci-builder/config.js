/* Armodilo CI Builder settings.
   Edit this file to add customers (importer emails) or part numbers (page 2 metal content).
   Customer keys are the company name in CAPITALS as it appears on the FedEx invoice.
   Part keys match the FedEx "Part Nbr" (capitals and spaces are ignored, so "ALUR - 14" matches "ALUR-14").

   breakdown_url - the Google Doc's Apps Script link (ends in /exec). When set, the page reads the
                   Material Breakdown Doc every time it opens; parts below are the fallback.

   Each part has:
     title       - text after "Line 1: " on page 2
     title_color - optional "black" (used for the plastic Tablet Fit Kit, which has no metal)
     components  - one block per metal component. Fields:
       name, note       optional: "Component: <name> (<note>)"
       hs, qty          optional: "H.S. Code:" and "Quantity:" lines
       material         "Aluminum" or "Steel"
       value            US$ number (multiplied by the line quantity)
       value_label      defaults to "Total Value" (steel uses "Total Component Value")
       weight           kg as text, e.g. "0.50" (multiplied by the line quantity)
       weight_label     defaults to "Total Weight" (steel uses "Total Component Weight")
       weight_suffix    optional text after "kg" (only shown when quantity is 1)
       origin           list of [label, country] lines
       reference        file name shown after "Reference:"
     item        - defaults used when you add this product as an extra line
                   (part number, HS code, unit value in US$, description) */
window.CONFIG = {
  "version": "2026-11-11a",
  "breakdown_url": "https://script.google.com/macros/s/AKfycbxeHnpHqSgH2aBnbZHycKP3XvCLJLLZuPnT7OODBty5QT4xEF5awCECXbJ2ypnV8ACj/exec",
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
  // Used when the FedEx file's Comments 1)-3) are blank (they can be changed on the page before downloading)
  "default_comments": [
    "*** Blanket USMCA on file under FTN ACC # 4869598 ***",
    "I hereby certify that the goods covered by this shipment qualifies as an originating good",
    "for the purposes of preferential tariff treatment under USMCS/T-MEC/CUSMA"
  ],
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
};
