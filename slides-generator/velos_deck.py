"""
Velos Logistics — case study deck (reference example).

Shows the end-to-end pattern this repo is built around:

  1. Think strategically with the user FIRST (in the chat).
     The thesis here = "Consortium Pivot to cross the Disruption Gate."
  2. Build slide content using the helpers in esade_design.py.
  3. Output three deliverables:
       - Velos_Full_Deck.pptx       (12 slides)
       - Velos_Presentation.pptx    (3 slides — title + diagnosis + recommendation)
       - Velos_Speaker_Scripts.pdf  (Presenter A / B split, Q&A prep)

Run with:
    python3 velos_deck.py
"""

import os

from esade_design import *
from reportlab.platypus import Paragraph, Spacer, PageBreak


# =============================================================================
# FULL DECK — 12 SLIDES
# =============================================================================

def build_full_deck(path):
    prs = setup_presentation()
    TOTAL = 13

    # ---- 1: TITLE ----------------------------------------------------------
    s = add_blank_slide(prs)
    add_rect(s, 0, 0, SLIDE_W, SLIDE_H, fill=NAVY)
    add_rect(s, 0, Inches(3.2), SLIDE_W, Inches(0.05), fill=ACCENT)
    add_text(s, Inches(0.8), Inches(2.2), Inches(11.7), Inches(1.0),
             "Velos Logistics", size=54, bold=True, color=WHITE)
    add_text(s, Inches(0.8), Inches(3.4), Inches(11.7), Inches(0.8),
             "From Box-Mover to Industry Standard",
             size=28, color=WHITE)
    add_text(s, Inches(0.8), Inches(4.3), Inches(11.7), Inches(0.6),
             "A Consortium Pivot to Cross the Disruption Gate",
             size=18, italic=True, color=GREY_LT)
    add_text(s, Inches(0.8), Inches(6.6), Inches(11.7), Inches(0.4),
             "Group X  ·  Technology & Digital Business  ·  Prof. Marc Sansó  ·  April 2026",
             size=11, color=GREY_LT)

    # ---- 2: EXECUTIVE SUMMARY ---------------------------------------------
    s = add_blank_slide(prs)
    slide_chrome(s, "Executive Summary — Anchor the Standard, Don't Own It",
                 "Reject the binary. The third path completes the Triple Crown.",
                 page_num=2, total=TOTAL)
    cols = [
        ("SITUATION",
         ["Velos runs two structurally opposed business models on one balance sheet.",
          ("Prime (Model 1.0): ", "linear cost structure, 10% gross margin, dying slowly."),
          ("Trust (Model 2.0 attempt): ", "83% gross margin, burning €15M, negligible adoption."),
          ("Stated ambition: ", "Model 3.0 — industry protocol via blockchain, smart contracts, single source of truth.")]),
        ("COMPLICATION",
         [("Velos is leaping Model 1.0 → 3.0 directly ", "— skipping the 2.0 ecosystem work that makes 3.0 viable."),
          "Triple Crown: ✓ Native Interface · ✓ Business Model · ✗ Dominant Ecosystem.",
          ("Open-Source Paradox: ",
           "solo orchestration in pre-standardization means Velos pays every standardization cost while application-layer rivals capture the future margins.")]),
        ("RESOLUTION",
         [("Open the protocol ", "as a neutral standard."),
          ("Anchor a consortium ", "of 2–3 competitors + 1 bank + 1 customs body — the 2.0 ecosystem foundation."),
          ("Gate Prime cash flow ", "against 18-month ecosystem milestones."),
          "Sequence the transformation: build 2.0 first, claim 3.0 once the ecosystem is self-sustaining."]),
    ]
    n = 3
    gap = Inches(0.3)
    col_w = (Inches(12.3) - gap * (n - 1)) / n
    for i, (title, body) in enumerate(cols):
        x = Inches(0.5) + (col_w + gap) * i
        add_rect(s, x, Inches(1.6), col_w, Inches(0.55), fill=NAVY)
        add_text(s, x + Inches(0.2), Inches(1.6), col_w - Inches(0.4), Inches(0.55),
                 title, size=13, bold=True, color=WHITE,
                 anchor=MSO_ANCHOR.MIDDLE)
        add_rect(s, x, Inches(2.15), col_w, Inches(4.5),
                 fill=BG_LIGHT, line=GREY_LT)
        add_bullets(s, x + Inches(0.2), Inches(2.30),
                    col_w - Inches(0.4), Inches(4.2),
                    body, size=12, line_spacing=1.3)

    add_callout(s, Inches(0.5), Inches(6.75), Inches(12.3), Inches(0.32),
                "Recommendation in one line: Velos must stop trying to own the protocol — and start curating the ecosystem that makes the protocol possible.",
                fill=ACCENT)

    # ---- 3: Q1 — TWO COMPANIES UNDER ONE ROOF ----------------------------
    s = add_blank_slide(prs)
    slide_chrome(s, "Q1 · Velos Operates Two Structurally Opposed Business Models",
                 "Applying the 6 Business Model Dimensions to Prime vs Trust",
                 page_num=3, total=TOTAL)
    data = [
        ["Dimension", "Velos Prime (Model 1.0)", "Velos Trust (Model 2.0 → 3.0 ambition)"],
        ["Monetization", "Per-shipment freight fees · transactional", "Subscription + transaction take-rate · recurring"],
        ["Cost Structure", "Heavily variable: fuel, drivers, fleet, leases", "Heavily fixed: R&D, dev, cloud — marginal cost ≈ 0"],
        ["Margin vs Volume", "Volume-driven · 10% gross margin", "Margin-driven · 83% gross margin"],
        ["Average Selling Price", "Per-shipment freight rate", "Subscription tier + small % per transaction"],
        ["Scalability", "Linear: revenue and cost scale 1:1", "Exponential post-breakeven · marginal cost ≈ 0"],
        ["Risk Profile", "Fuel · recession · fleet utilization · commoditization", "Adoption · ecosystem dependence · runway · Open-Source Paradox"],
    ]
    add_table(s, Inches(0.5), Inches(1.55), Inches(12.3), Inches(4.7),
              data, first_col_bold=True, first_col_fill=NAVY_LITE,
              col_widths=[Inches(2.0), Inches(4.7), Inches(5.6)],
              size=11, header_size=12)
    add_callout(s, Inches(0.5), Inches(6.45), Inches(12.3), Inches(0.55),
                "These are not two divisions of one company — they are two different companies that share a balance sheet.",
                fill=ACCENT, size=14, italic=True)

    # ---- 4: Q1 cont — OPERATING LEVERAGE ---------------------------------
    s = add_blank_slide(prs)
    slide_chrome(s, "Q1 cont · Cost Structure Asymmetry Is the Strategic Story",
                 "Operating leverage: Prime's profit is bounded; Trust's is not.",
                 page_num=4, total=TOTAL)

    add_text(s, Inches(0.5), Inches(1.5), Inches(6.0), Inches(0.4),
             "PRIME — Service Model (perfectly variable)",
             size=14, bold=True, color=NAVY)
    cl, ct, cw, ch = (Inches(0.7), Inches(2.0), Inches(5.6), Inches(3.5))
    add_rect(s, cl, ct + ch, cw, Emu(8000), fill=GREY)
    add_rect(s, cl, ct, Emu(8000), ch, fill=GREY)
    rev = s.shapes.add_connector(1, cl, ct + ch, cl + cw, ct + Inches(0.3))
    rev.line.color.rgb = NAVY
    rev.line.width = Pt(3)
    cost = s.shapes.add_connector(1, cl, ct + ch - Inches(0.4),
                                  cl + cw, ct + Inches(0.7))
    cost.line.color.rgb = AMBER
    cost.line.width = Pt(2.5)
    add_text(s, cl + cw - Inches(1.4), ct + Inches(0.0), Inches(1.4), Inches(0.3),
             "Revenue", size=10, bold=True, color=NAVY)
    add_text(s, cl + cw - Inches(1.4), ct + Inches(0.45), Inches(1.4), Inches(0.3),
             "Cost", size=10, bold=True, color=AMBER)
    add_text(s, cl, ct + ch + Inches(0.05), cw, Inches(0.3),
             "Volume →", size=9, color=GREY, align=PP_ALIGN.CENTER)
    add_text(s, cl - Inches(0.05), ct - Inches(0.3), cw, Inches(0.3),
             "$  Constant margin · no leverage", size=10, italic=True, color=GREY)

    add_text(s, Inches(7.0), Inches(1.5), Inches(6.0), Inches(0.4),
             "TRUST — SaaS Model (high fixed, ~0 marginal)",
             size=14, bold=True, color=NAVY)
    cl2, ct2, cw2, ch2 = (Inches(7.2), Inches(2.0), Inches(5.6), Inches(3.5))
    add_rect(s, cl2, ct2 + ch2, cw2, Emu(8000), fill=GREY)
    add_rect(s, cl2, ct2, Emu(8000), ch2, fill=GREY)
    fc = s.shapes.add_connector(1, cl2, ct2 + Inches(2.3),
                                cl2 + cw2, ct2 + Inches(2.3))
    fc.line.color.rgb = AMBER
    fc.line.width = Pt(2.5)
    rev2 = s.shapes.add_connector(1, cl2, ct2 + ch2,
                                  cl2 + cw2, ct2 + Inches(0.1))
    rev2.line.color.rgb = NAVY
    rev2.line.width = Pt(3)
    be = s.shapes.add_shape(MSO_SHAPE.OVAL,
                            cl2 + Inches(2.4), ct2 + Inches(2.15),
                            Inches(0.25), Inches(0.25))
    fill_shape(be, ACCENT)
    add_text(s, cl2 + Inches(2.6), ct2 + Inches(2.0), Inches(2.0), Inches(0.3),
             "Breakeven", size=9, bold=True, color=ACCENT)
    add_text(s, cl2 + cw2 - Inches(1.6), ct2 + Inches(0.0), Inches(1.6), Inches(0.3),
             "Revenue", size=10, bold=True, color=NAVY)
    add_text(s, cl2 + Inches(0.1), ct2 + Inches(2.05), Inches(2.0), Inches(0.3),
             "Fixed costs", size=10, bold=True, color=AMBER)
    add_text(s, cl2, ct2 + ch2 + Inches(0.05), cw2, Inches(0.3),
             "Volume →", size=9, color=GREY, align=PP_ALIGN.CENTER)
    add_text(s, cl2 - Inches(0.05), ct2 - Inches(0.3), cw2, Inches(0.3),
             "$  Expanding margin · exponential leverage", size=10, italic=True, color=GREY)

    add_text(s, Inches(0.5), Inches(5.85), Inches(6.0), Inches(0.4),
             "$2X revenue → $2X profit  (capped)",
             size=13, bold=True, color=AMBER, align=PP_ALIGN.CENTER)
    add_text(s, Inches(7.0), Inches(5.85), Inches(6.0), Inches(0.4),
             "$2X revenue → $6X+ profit  (unbounded)",
             size=13, bold=True, color=GREEN, align=PP_ALIGN.CENTER)

    add_callout(s, Inches(0.5), Inches(6.4), Inches(12.3), Inches(0.6),
                "Trust's 83% gross margin is meaningless without volume. The unit economics demand scale that requires an ecosystem Velos doesn't yet have.",
                fill=ACCENT, size=13)

    # ---- 5: Q2 — MODEL 1.0 vs 2.0 ----------------------------------------
    s = add_blank_slide(prs)
    slide_chrome(s, "Q2 · Prime Was Built for Linear Value. Trust Demands Network Value.",
                 "The Evolutionary Matrix exposes the structural mismatch.",
                 page_num=5, total=TOTAL)
    data = [
        ["", "Velos Prime (Model 1.0)", "Velos Trust (Model 2.0 → 3.0 ambition)"],
        ["Defining logic", "Product/industry-centric", "Ecosystem-centric"],
        ["Architecture", "Independent units · additive growth", "Interdependent layers · asymmetric growth"],
        ["Value creation", "Move boxes A → B (linear value chain)", "Orchestrate matching, trust, payments between independent agents"],
        ["Moat", "Operational efficiency (fragile, easily commoditized)", "Shared infrastructure + switching costs + network effects"],
        ["Growth math", "Value = A + B + C", "Value > sum of parts"],
    ]
    add_table(s, Inches(0.5), Inches(1.55), Inches(12.3), Inches(4.4),
              data, first_col_bold=True, first_col_fill=NAVY_LITE,
              col_widths=[Inches(2.3), Inches(4.6), Inches(5.4)],
              size=11.5, header_size=12)
    add_callout(s, Inches(0.5), Inches(6.2), Inches(12.3), Inches(0.8),
                "Velos is attempting Model 2.0 economics with a Model 1.0 mindset, while reaching for Model 3.0 outcomes alone. That is why adoption is stalled.",
                fill=ACCENT, size=13)

    # ---- 6: Q2 cont — MODEL 3.0 AMBITION + THE LEAP PROBLEM --------------
    s = add_blank_slide(prs)
    slide_chrome(s, "Q2 cont · The Real Endgame Is Model 3.0 — and That's Why the Strategy Stalls",
                 "Trust isn't just a platform play. Velos's stated ambition is to become the protocol layer of logistics.",
                 page_num=6, total=TOTAL)

    stages = [
        ("MODEL 1.0", "LINEAR CHAIN",
         "Velos Prime today",
         "Move boxes A → B · 10% margin",
         NAVY),
        ("MODEL 2.0", "PLATFORM",
         "Velos Trust (attempt)",
         "Orchestrate · 83% margin · no volume",
         AMBER),
        ("MODEL 3.0", "PROTOCOL",
         "Velos's stated ambition",
         "Define the rules others build on",
         ACCENT),
    ]
    card_top = Inches(1.55)
    card_h = Inches(1.85)
    arrow_w = Inches(0.3)
    arrow_gap = Inches(0.1)
    card_w = (Inches(12.3) - (arrow_w + arrow_gap * 2) * 2) / 3
    cursor_x = Inches(0.5)
    for i, (model, kind, who, desc, color) in enumerate(stages):
        x = cursor_x
        add_rect(s, x, card_top, card_w, Inches(0.45), fill=color)
        add_text(s, x, card_top, card_w, Inches(0.45),
                 model, size=13, bold=True, color=WHITE,
                 align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
        add_rect(s, x, card_top + Inches(0.45), card_w, card_h - Inches(0.45),
                 fill=BG_LIGHT, line=GREY_LT)
        add_text(s, x, card_top + Inches(0.55), card_w, Inches(0.45),
                 kind, size=15, bold=True, color=color,
                 align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
        add_text(s, x + Inches(0.15), card_top + Inches(1.05),
                 card_w - Inches(0.3), Inches(0.35),
                 who, size=11, italic=True, color=NAVY,
                 align=PP_ALIGN.CENTER)
        add_text(s, x + Inches(0.15), card_top + Inches(1.4),
                 card_w - Inches(0.3), Inches(0.4),
                 desc, size=10, color=GREY,
                 align=PP_ALIGN.CENTER)
        cursor_x = x + card_w
        if i < 2:
            arr_x = cursor_x + arrow_gap
            arr_y = card_top + (card_h - Inches(0.32)) / 2
            arr = s.shapes.add_shape(MSO_SHAPE.RIGHT_ARROW,
                                     arr_x, arr_y, arrow_w, Inches(0.32))
            fill_shape(arr, GREY)
            cursor_x = arr_x + arrow_w + arrow_gap

    mid_top = Inches(3.6)
    col_w_m = Inches(6.0)
    col_gap_m = Inches(0.3)

    add_rect(s, Inches(0.5), mid_top, col_w_m, Inches(2.6),
             fill=BG_LIGHT, line=GREY_LT)
    add_rect(s, Inches(0.5), mid_top, col_w_m, Inches(0.5), fill=NAVY)
    add_text(s, Inches(0.5), mid_top, col_w_m, Inches(0.5),
             "Why Trust is reaching for Model 3.0 (not just 2.0)",
             size=12, bold=True, color=WHITE,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    add_bullets(s, Inches(0.7), mid_top + Inches(0.65),
                col_w_m - Inches(0.4), Inches(1.85),
                [("Smart contracts: ", "autonomous execution of payments and compliance — no human intermediary."),
                 ("Single source of truth: ", "unified data layer across shippers, carriers, banks, customs."),
                 ("Industry-wide rules: ", "Velos sets the protocol all participants must operate within.")],
                size=11, line_spacing=1.3)

    right_x = Inches(0.5) + col_w_m + col_gap_m
    add_rect(s, right_x, mid_top, col_w_m, Inches(2.6),
             fill=BG_LIGHT, line=GREY_LT)
    add_rect(s, right_x, mid_top, col_w_m, Inches(0.5), fill=ACCENT)
    add_text(s, right_x, mid_top, col_w_m, Inches(0.5),
             "The leap problem · 1.0 → 3.0 skipping 2.0",
             size=12, bold=True, color=WHITE,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    add_bullets(s, right_x + Inches(0.2), mid_top + Inches(0.65),
                col_w_m - Inches(0.4), Inches(1.85),
                [("No ecosystem: ", "Trust lacks critical mass — network effects haven't ignited."),
                 ("No standard: ", "no industry-wide adoption of Velos's protocol."),
                 ("No neutral trust: ", "competitors will not adopt infrastructure owned by a rival.")],
                size=11, line_spacing=1.3)

    add_callout(s, Inches(0.5), Inches(6.45), Inches(12.3), Inches(0.55),
                "No ecosystem → no platform → no protocol. Velos is trying to define the rules of the system before building the system itself.",
                fill=ACCENT, size=14)

    # ---- 7: Q2 cont — THREE STRUCTURAL REASONS ---------------------------
    s = add_blank_slide(prs)
    slide_chrome(s, "Q2 cont · Three Structural Reasons the 1.0 → 3.0 Leap Cannot Work Alone",
                 "Single-firm platform plays fail in pre-standardization markets.",
                 page_num=7, total=TOTAL)
    columns = [
        ("Chicken-and-egg unsolved",
         ["Trust requires banks, customs, and competitors to join in parallel.",
          "Each waits for the others.",
          "No critical interaction density → no network effects → no platform."]),
        ("Ownership ≠ orchestration",
         ["Velos behaves as proprietor. Model 2.0 demands acting as host.",
          "Competitors will not join an ecosystem owned by a competitor.",
          "Governance must feel neutral, not extractive."]),
        ("No standards yet",
         ["Every counterparty requires a custom integration with Velos's proprietary protocol.",
          "CAC explodes — every onboarding is bespoke.",
          "This is the CFO's 'mud roads' problem."]),
    ]
    draw_three_columns(s, Inches(0.5), Inches(1.7), Inches(12.3), Inches(4.7),
                       columns)
    add_callout(s, Inches(0.5), Inches(6.55), Inches(12.3), Inches(0.55),
                "Model 2.0 platforms cannot be willed into existence by a single firm. They emerge from coordinated multi-party investment in shared infrastructure.",
                fill=NAVY, size=13)

    # ---- 8: Q3 — INEVITABILITY CASCADE -----------------------------------
    s = add_blank_slide(prs)
    slide_chrome(s, "Q3 · Trust Has Cleared Phases 1–3. Phase 4 Is the Cliff.",
                 "The Inevitability Cascade locates exactly where Velos is stranded.",
                 page_num=8, total=TOTAL)
    statuses = ["ok", "ok", "warn", "fail"]
    labels = ["Phase 1\nTech Functionality", "Phase 2\nHard Use Cases",
              "Phase 3\nBusiness Model", "Phase 4\nDominant Ecosystem"]
    sublabels = [
        "DLT works · smart contracts execute reliably",
        "20% admin cost reduction is real, repeatable",
        "83% gross margin proves logic; volume missing",
        "No standards · no orchestrator · no self-propulsion",
    ]
    draw_inevitability_cascade(s, Inches(0.5), Inches(1.8), Inches(12.3), Inches(2.0),
                               statuses, labels, sublabels, highlight_last=True)
    add_text(s, Inches(0.5), Inches(4.05), Inches(12.3), Inches(0.4),
             "↑  Phase 4 is the Disruption Gate — the only phase where structural disruption actually occurs.",
             size=14, bold=True, color=ACCENT, align=PP_ALIGN.CENTER)
    add_rect(s, Inches(0.5), Inches(4.8), Inches(12.3), Inches(1.9),
             fill=BG_LIGHT, line=GREY_LT)
    add_text(s, Inches(0.7), Inches(4.95), Inches(11.9), Inches(0.4),
             "Strategic implication",
             size=13, bold=True, color=NAVY)
    add_bullets(s, Inches(0.7), Inches(5.35), Inches(11.9), Inches(1.3),
                ["Without crossing Phase 4, the first three phases are stranded investments.",
                 "Velos has built the engine; the road network does not exist yet.",
                 "Solo investment in Phase 4 is structurally a losing bet — standards emerge from multi-party governance."],
                size=12)

    # ---- 9: Q3 cont — KINETIC ASYMMETRY ----------------------------------
    s = add_blank_slide(prs)
    slide_chrome(s, "Q3 cont · The 'Trough of Disillusionment' Is a Kinetic Asymmetry",
                 "Layers 1–3 raced ahead. Layer 4 hasn't started.",
                 page_num=9, total=TOTAL)
    layers = [
        ("Layer 1 — Tech functionality", 0.95, NAVY),
        ("Layer 2 — Hard use cases",     0.70, NAVY_LITE),
        ("Layer 3 — Business model",     0.55, AMBER),
        ("Layer 4 — Dominant ecosystem", 0.15, ACCENT),
    ]
    draw_layer_arrows(s, Inches(0.5), Inches(1.6), Inches(12.3), Inches(3.2), layers)
    add_rect(s, Inches(0.5), Inches(5.0), Inches(12.3), Inches(1.9),
             fill=BG_LIGHT, line=GREY_LT)
    add_text(s, Inches(0.7), Inches(5.15), Inches(11.9), Inches(0.4),
             "What this actually means",
             size=13, bold=True, color=NAVY)
    add_bullets(s, Inches(0.7), Inches(5.55), Inches(11.9), Inches(1.3),
                ["The CTO's 'Trough of Disillusionment' is not a hype crash — it is the predictable consequence of Layer 4 lagging Layers 1–3.",
                 "Banks, customs authorities, ERP vendors, and competitor logistics firms are the rate limit — not the technology.",
                 "The fix is not more code. It is multi-party governance."],
                size=12)

    # ---- 10: Q3 cont — OPEN-SOURCE PARADOX -------------------------------
    s = add_blank_slide(prs)
    slide_chrome(s, "Q3 cont · Solo Orchestration Walks Velos Into the Open-Source Paradox",
                 "The firm that pays for the standard rarely captures its profits.",
                 page_num=10, total=TOTAL)

    box_l = (Inches(0.5), Inches(1.6), Inches(6.0), Inches(4.7))
    add_rect(s, *box_l, fill=BG_LIGHT, line=ACCENT)
    add_rect(s, box_l[0], box_l[1], box_l[2], Inches(0.55), fill=ACCENT)
    add_text(s, box_l[0] + Inches(0.2), box_l[1], box_l[2] - Inches(0.4), Inches(0.55),
             "SOLO PATH (current trajectory)",
             size=13, bold=True, color=WHITE, anchor=MSO_ANCHOR.MIDDLE)
    add_bullets(s, box_l[0] + Inches(0.25), box_l[1] + Inches(0.75),
                box_l[2] - Inches(0.5), box_l[3] - Inches(0.85),
                ["Velos absorbs €15M+ in standardization costs alone.",
                 "Each new counterparty is a custom integration → CAC keeps rising.",
                 "When industry-wide standards eventually emerge anyway, application-layer competitors enter cheaply.",
                 ("Outcome: ", "Velos becomes the unpaid infrastructure for someone else's business model.")],
                size=12, line_spacing=1.3)

    box_r = (Inches(6.8), Inches(1.6), Inches(6.0), Inches(4.7))
    add_rect(s, *box_r, fill=BG_LIGHT, line=GREEN)
    add_rect(s, box_r[0], box_r[1], box_r[2], Inches(0.55), fill=GREEN)
    add_text(s, box_r[0] + Inches(0.2), box_r[1], box_r[2] - Inches(0.4), Inches(0.55),
             "CONSORTIUM PATH (recommended)",
             size=13, bold=True, color=WHITE, anchor=MSO_ANCHOR.MIDDLE)
    add_bullets(s, box_r[0] + Inches(0.25), box_r[1] + Inches(0.75),
                box_r[2] - Inches(0.5), box_r[3] - Inches(0.85),
                ["Standardization costs shared across founding members.",
                 "Onboarding becomes self-service via the open standard.",
                 "Velos anchors governance and sets the founding rules.",
                 ("Outcome: ", "Velos captures premium positioning as standard curator + reference implementation.")],
                size=12, line_spacing=1.3)

    add_callout(s, Inches(0.5), Inches(6.45), Inches(12.3), Inches(0.55),
                "In B2B ecosystems, the firm that pays for standardization rarely captures its profits — unless it builds governance leverage early.",
                fill=NAVY, size=13)

    # ---- 11: Q4 — REJECT THE BINARY --------------------------------------
    s = add_blank_slide(prs)
    slide_chrome(s, "Q4 · Both Board Options Are Losing Moves — A Third Path Exists",
                 "Comparing Retrench, Solo Double Down, and the Consortium Pivot.",
                 page_num=11, total=TOTAL)
    data = [
        ["", "Retrench", "Solo Double Down", "Consortium Pivot ★"],
        ["Time horizon", "Dies slowly (~10 yrs)", "Cash crisis in ~2 yrs", "Standards emerge in 3–5 yrs"],
        ["Primary risk", "Strategic obsolescence", "Solvency + Open-Source Paradox", "Governance complexity, slower ramp"],
        ["Triple Crown progress", "Abandons all three", "Stalled at Interface + BM", "Completes the third (Ecosystem)"],
        ["Profit pool capture", "None — commoditized utility", "Infrastructure layer at best", "Anchor + standard curator"],
        ["Optionality", "Low", "Low", "High — spin out, license, or scale"],
    ]
    add_table(s, Inches(0.5), Inches(1.55), Inches(12.3), Inches(4.4),
              data, first_col_bold=True, first_col_fill=NAVY_LITE,
              col_widths=[Inches(2.5), Inches(2.8), Inches(3.2), Inches(3.8)],
              size=11.5, header_size=12)
    add_callout(s, Inches(0.5), Inches(6.2), Inches(12.3), Inches(0.8),
                "Retrench and Double Down both accept defeat — one slow, one fast. Only the Consortium Pivot completes the Triple Crown.",
                fill=ACCENT, size=14)

    # ---- 12: Q4 — CONSORTIUM PIVOT PLAYBOOK ------------------------------
    s = add_blank_slide(prs)
    slide_chrome(s, "Q4 · The Consortium Pivot Playbook — Three Moves",
                 "Convert competitors into co-founders of a shared standard.",
                 page_num=12, total=TOTAL)
    columns = [
        ("OPEN THE PROTOCOL",
         ["Release core data formats as an open standard under neutral governance.",
          "Converts proprietary lock-in (which is blocking adoption) into ecosystem gravity.",
          "Velos retains the reference-implementation premium."]),
        ("ANCHOR THE CONSORTIUM",
         ["Recruit 2–3 logistics competitors + 1 anchor bank + 1 customs/regulator as founding members.",
          "Solves chicken-and-egg by guaranteeing minimum interaction density on day one.",
          "Activates Network Bridging defense (Quadrant D): synergies across previously separate networks."]),
        ("GATE THE RUNWAY",
         ["Prime cash funds Trust against hard 18-month milestones: ARR, consortium signatures, integration depth.",
          "If milestones miss → structured spin-out of Trust into the consortium; Velos harvests Prime.",
          "Caps downside without forcing a binary today."]),
    ]
    draw_three_columns(s, Inches(0.5), Inches(1.7), Inches(12.3), Inches(4.7),
                       columns)
    add_callout(s, Inches(0.5), Inches(6.55), Inches(12.3), Inches(0.55),
                "Each move is reversible. Together they build the Model 2.0 ecosystem that makes Model 3.0 viable — sequential, not simultaneous.",
                fill=NAVY, size=13)

    # ---- 13: Q4 — RECOMMENDATION TIMELINE --------------------------------
    s = add_blank_slide(prs)
    slide_chrome(s, "Q4 · Recommendation — Anchor the Standard. Don't Own It.",
                 "An 18-month gated path with explicit kill criteria preserves optionality.",
                 page_num=13, total=TOTAL)

    bar_top = Inches(2.4)
    bar_left = Inches(1.5)
    bar_width = Inches(10.3)
    bar_h = Inches(0.18)
    node_d = Inches(0.5)
    add_rect(s, bar_left, bar_top + Inches(0.16), bar_width, bar_h, fill=GREY_LT)

    milestones = [
        ("Month 0–12", "Open protocol · recruit consortium · target 3 founding members",
         0.0, "left",  NAVY),
        ("Month 12 GATE", "≥2 consortium members signed + ≥€5M ARR commitments",
         0.5, "center", ACCENT),
        ("Month 18 GATE", "Standard ratified · first cross-firm live transaction",
         1.0, "right", ACCENT),
    ]
    for label, desc, pos, anchor, color in milestones:
        cx = bar_left + Emu(int(bar_width * pos))
        nx = cx - node_d / 2
        ny = bar_top + bar_h / 2 - node_d / 2
        node = s.shapes.add_shape(MSO_SHAPE.OVAL, nx, ny, node_d, node_d)
        fill_shape(node, color)

        block_w = Inches(3.1)
        if anchor == "left":
            tx = bar_left - Inches(0.3)
            text_align = PP_ALIGN.LEFT
        elif anchor == "right":
            tx = bar_left + bar_width - block_w + Inches(0.3)
            text_align = PP_ALIGN.RIGHT
        else:
            tx = cx - block_w / 2
            text_align = PP_ALIGN.CENTER

        add_text(s, tx, bar_top - Inches(0.7), block_w, Inches(0.4),
                 label, size=13, bold=True, color=color,
                 align=text_align)
        add_text(s, tx, bar_top + Inches(0.85), block_w, Inches(0.9),
                 desc, size=11, color=GREY_DK,
                 align=text_align)

    add_rect(s, Inches(0.5), Inches(4.5), Inches(12.3), Inches(1.5),
             fill=BG_LIGHT, line=GREY_LT)
    add_text(s, Inches(0.7), Inches(4.6), Inches(11.9), Inches(0.4),
             "If any gate misses",
             size=13, bold=True, color=ACCENT)
    add_bullets(s, Inches(0.7), Inches(5.0), Inches(11.9), Inches(1.0),
                ["Structured spin-out of Trust into the consortium (or to a strategic acquirer / PE).",
                 "Prime returns to optimization mode with proceeds — capital recycled, optionality preserved."],
                size=12)

    add_callout(s, Inches(0.5), Inches(6.25), Inches(12.3), Inches(0.75),
                "The freight market is dying slowly. The platform market is not yet born. The Consortium Pivot is how Velos buys the right to be born.",
                fill=ACCENT, size=14)

    prs.save(path)
    print(f"  Saved: {path}")


# =============================================================================
# PRESENTATION DECK — 3 SLIDES
# =============================================================================

def build_presentation_deck(path):
    prs = setup_presentation()

    # ---- 1: TITLE ----------------------------------------------------------
    s = add_blank_slide(prs)
    add_rect(s, 0, 0, SLIDE_W, SLIDE_H, fill=NAVY)
    add_rect(s, 0, Inches(3.2), SLIDE_W, Inches(0.05), fill=ACCENT)
    add_text(s, Inches(0.8), Inches(2.2), Inches(11.7), Inches(1.0),
             "Velos Logistics", size=54, bold=True, color=WHITE)
    add_text(s, Inches(0.8), Inches(3.4), Inches(11.7), Inches(0.8),
             "From Box-Mover to Industry Standard",
             size=28, color=WHITE)
    add_text(s, Inches(0.8), Inches(4.3), Inches(11.7), Inches(0.6),
             "A Consortium Pivot to Cross the Disruption Gate",
             size=18, italic=True, color=GREY_LT)
    add_text(s, Inches(0.8), Inches(6.5), Inches(11.7), Inches(0.4),
             "Group X  ·  [Presenter A] & [Presenter B]",
             size=12, color=GREY_LT)
    add_text(s, Inches(0.8), Inches(6.85), Inches(11.7), Inches(0.4),
             "Technology & Digital Business  ·  Prof. Marc Sansó  ·  April 2026",
             size=11, color=GREY_LT)

    # ---- 2: DIAGNOSIS -----------------------------------------------------
    s = add_blank_slide(prs)
    slide_chrome(s, "Diagnosis · Velos Is One Phase Short of Disruption")

    add_text(s, Inches(0.5), Inches(1.4), Inches(6.0), Inches(0.4),
             "Two structurally opposed business models",
             size=13, bold=True, color=NAVY)
    mini = [
        ["", "Prime (1.0)", "Trust (2.0 → 3.0 reach)"],
        ["Cost structure", "Variable · linear", "Fixed · exponential"],
        ["Gross margin", "10%", "83%"],
        ["Scalability", "Capped", "Unbounded post-volume"],
    ]
    add_table(s, Inches(0.5), Inches(1.85), Inches(6.0), Inches(2.0),
              mini, first_col_bold=True, first_col_fill=NAVY_LITE,
              col_widths=[Inches(1.7), Inches(2.0), Inches(2.3)],
              size=11, header_size=11.5)

    add_text(s, Inches(0.5), Inches(4.1), Inches(6.0), Inches(0.4),
             "Inevitability Cascade — Velos is stranded at Phase 4",
             size=13, bold=True, color=NAVY)
    statuses = ["ok", "ok", "warn", "fail"]
    labels = ["Tech", "Use Cases", "Business Model", "Ecosystem"]
    draw_inevitability_cascade(s, Inches(0.5), Inches(4.55), Inches(6.0), Inches(1.1),
                               statuses, labels, highlight_last=True, compact=True)
    add_text(s, Inches(0.5), Inches(5.75), Inches(6.0), Inches(0.4),
             "↑ Phase 4 = the Disruption Gate",
             size=11, bold=True, italic=True, color=ACCENT,
             align=PP_ALIGN.CENTER)

    add_text(s, Inches(7.0), Inches(1.4), Inches(5.8), Inches(0.4),
             "Triple Crown audit",
             size=13, bold=True, color=NAVY)
    items = [
        ("Native Interface — DLT works", "ok"),
        ("Dominant Business Model — 83% gross margin", "ok"),
        ("Dominant Ecosystem — no standards, no orchestrator", "fail"),
    ]
    draw_triple_crown(s, Inches(7.0), Inches(1.95), Inches(5.8), Inches(2.4), items)

    add_rect(s, Inches(7.0), Inches(4.7), Inches(5.8), Inches(2.3), fill=ACCENT)
    add_text(s, Inches(7.2), Inches(4.85), Inches(5.4), Inches(0.4),
             "THE TRAP",
             size=12, bold=True, color=WHITE)
    add_text(s, Inches(7.2), Inches(5.25), Inches(5.4), Inches(1.7),
             "Velos is leaping 1.0 → 3.0 alone, skipping the 2.0 ecosystem work. Open-Source Paradox: Velos pays for the standard, application-layer rivals capture the margins.",
             size=13, italic=True, color=WHITE, anchor=MSO_ANCHOR.TOP)

    # ---- 3: RECOMMENDATION ------------------------------------------------
    s = add_blank_slide(prs)
    slide_chrome(s, "Recommendation · Anchor the Standard. Don't Own It.")

    columns = [
        ("OPEN PROTOCOL",
         ["Release data formats as open standard.",
          "Convert proprietary lock-in into ecosystem gravity.",
          "Retain reference-implementation premium."]),
        ("ANCHOR CONSORTIUM",
         ["2–3 competitors + 1 bank + 1 customs body as co-founders.",
          "Solves chicken-and-egg.",
          "Activates Network Bridging defense (Quadrant D)."]),
        ("GATE RUNWAY",
         ["Prime cash funds Trust on 18-month milestones (ARR · signatures · ratification).",
          "Spin out if missed.",
          "Caps downside, preserves optionality."]),
    ]
    draw_three_columns(s, Inches(0.5), Inches(1.4), Inches(12.3), Inches(3.6),
                       columns)

    data = [
        ["Retrench", "Solo Double Down", "Consortium Pivot ★"],
        ["Dies in ~10 years", "Dies in ~2 years", "Only path that completes the Triple Crown"],
        ["Abandons all three phases", "Stranded at Phase 3", "Crosses the Disruption Gate"],
    ]
    add_table(s, Inches(0.5), Inches(5.25), Inches(12.3), Inches(1.55),
              data, first_col_bold=False,
              col_widths=[Inches(3.5), Inches(3.8), Inches(5.0)],
              size=12, header_size=12)

    add_callout(s, Inches(0.5), Inches(6.9), Inches(12.3), Inches(0.4),
                "Velos must stop trying to own the platform — and start curating the standard others build on.",
                fill=ACCENT, size=13)

    prs.save(path)
    print(f"  Saved: {path}")


# =============================================================================
# SPEAKER SCRIPTS PDF
# =============================================================================

def build_speaker_pdf(path):
    doc = make_pdf_doc(
        path,
        title="Velos Logistics — Speaker Scripts",
        author="Group X — Technology & Digital Business",
    )
    st = pdf_styles()
    story = []

    story.append(Paragraph("Velos Logistics — Speaker Scripts", st["title"]))
    story.append(Paragraph(
        "10-minute case defense · Technology & Digital Business · Prof. Marc Sansó · April 2026",
        st["subtitle"]
    ))

    story.append(Paragraph("Presenter assignment", st["section"]))
    story.append(Paragraph(
        "<b>Presenter A</b> opens (Slide 1) and delivers the diagnosis (Slide 2). "
        "<b>Presenter B</b> delivers the recommendation (Slide 3) and fields the first Q&amp;A. "
        "Either presenter can take questions inside their own area.",
        st["body"]
    ))

    story.append(Paragraph("Slide 1 — Title", st["section"]))
    story.append(Paragraph("Presenter A · ~30 seconds", st["meta"]))
    story.append(Paragraph(
        "“Good morning. Our argument is that the Velos board is being offered a false binary "
        "— Retrench or Double Down — and that both are losing moves. We will recommend a third "
        "path: a Consortium Pivot that converts Velos&#8217;s competitors into co-founders of an "
        "industry standard. We&#8217;ll diagnose why the platform is stalled in two minutes, then walk "
        "through the three moves that get Velos across the Disruption Gate.”",
        st["quote"]
    ))

    story.append(Paragraph("Slide 2 — Diagnosis", st["section"]))
    story.append(Paragraph("Presenter A · ~4 minutes", st["meta"]))
    diag = [
        "“First, the structural diagnosis.",
        "Velos operates two completely opposed business models on one balance sheet. Prime is a Model 1.0 "
        "freight business — heavily variable cost structure, 10% gross margin, scaling linearly with fuel "
        "and headcount. Trust attempts Model 2.0 — heavily fixed cost structure, 83% gross margin, "
        "exponential operating leverage past breakeven. These are not two divisions of one company. "
        "They are two different companies that share a balance sheet.",
        "Now apply the Inevitability Cascade. Trust has cleared three of the four phases. Phase 1, Tech "
        "Functionality — DLT works, smart contracts execute reliably. Phase 2, Hard Use Cases — the "
        "20% administrative cost reduction is real, measurable, and repeatable. Phase 3, Business Model "
        "— the 83% gross margin proves the unit economics work in principle. But Phase 4, Dominant "
        "Ecosystem, is missing entirely. No standards. No orchestrator. No network self-propulsion.",
        "Phase 4 is the Disruption Gate. It is the only phase where structural disruption actually "
        "happens. Without crossing it, the first three phases are stranded investments. Velos has built "
        "the engine. The road network does not exist yet.",
        "Apply the Triple Crown lens and the picture sharpens. Native Interface — yes. Dominant "
        "Business Model logic — yes. Dominant Ecosystem — no. That is why the CTO&#8217;s &#8216;Trough of "
        "Disillusionment&#8217; is not a tech failure. It is a kinetic asymmetry: Layer 1 raced ahead while "
        "Layer 4 hasn&#8217;t started. Banks, customs authorities, ERP vendors, and competitors are the rate "
        "limit — not the technology.",
        "There is a deeper structural read here, and it is the one our team converged on. Velos Prime is "
        "Model 1.0 — a linear value chain. Velos Trust today operates as Model 2.0 — a platform "
        "connecting shippers, carriers, banks, and customs. But Velos&#8217;s stated ambition — smart "
        "contracts that automate execution, a single source of truth across the industry, a protocol "
        "that defines how participants must operate — that is Model 3.0. The operating system of "
        "logistics. The trouble is structural: you cannot have Model 3.0 without first building "
        "Model 2.0. No ecosystem, no platform. No platform, no protocol. Velos is trying to define "
        "the rules of the system before the system itself exists.",
        "On the current path, this trap gets worse. If Velos keeps absorbing standardization costs "
        "alone, they walk straight into the Open-Source Paradox. When standards eventually emerge "
        "industry-wide, application-layer competitors enter cheaply and capture the margins. Velos "
        "becomes the unpaid infrastructure for someone else&#8217;s business model. That is the worst "
        "possible outcome — and it is exactly where the current strategy leads.”",
    ]
    for p in diag:
        story.append(Paragraph(p, st["quote"]))
    story.append(Paragraph(
        "<i>Hand-off:</i> &#8220;[Teammate name] will walk us through the recommendation.&#8221;",
        st["body"]
    ))

    story.append(PageBreak())

    story.append(Paragraph("Slide 3 — Recommendation", st["section"]))
    story.append(Paragraph("Presenter B · ~4 minutes", st["meta"]))
    rec = [
        "“Our recommendation is to reject both options the board is debating. Retrench means dying "
        "slowly in a commoditizing 1.0 market — the freight market is dying anyway, just over ten "
        "years instead of two. Solo Double Down accelerates the death and hands the future to whoever "
        "benefits from the standard Velos paid to create.",
        "The Consortium Pivot is the only move that completes the Triple Crown. Three concrete actions.",
        "<b>One: open the protocol.</b> Release the core data formats as an open standard under neutral "
        "governance. This sounds counterintuitive — why give it away? Because in B2B ecosystems, "
        "proprietary lock-in is exactly what blocks adoption. Opening the protocol converts CAC into "
        "ecosystem gravity, removes the single biggest objection competitors have, and Velos retains "
        "the premium of being the reference implementation.",
        "<b>Two: anchor the consortium.</b> Recruit two or three logistics competitors, one anchor bank, "
        "and one customs or regulatory body as co-founders of a governance body. This solves the "
        "chicken-and-egg problem on day one by guaranteeing minimum interaction density. It also "
        "activates Network Bridging — the strongest defense against multi-homing — by combining "
        "previously separate networks into a single value proposition no single platform can match.",
        "<b>Three: gate the runway.</b> Prime cash flow does not subsidize Trust indefinitely. It funds "
        "Trust against hard 18-month milestones — consortium signatures, ARR commitments, standard "
        "ratification, and first cross-firm live transaction. If milestones miss, Trust spins out into "
        "the consortium, Velos harvests Prime, and capital is returned. That caps the downside without "
        "forcing a binary today.",
        "The risk-adjusted picture is clear. Retrench has low risk and zero upside. Solo Double Down "
        "has high risk and a low probability of capture. The Consortium Pivot has moderate execution "
        "risk, but it is the only option that preserves optionality and aligns with how B2B disruption "
        "actually crystallizes — slow, structural, multi-party.",
        "Elena&#8217;s framing was right: dying slowly in ten years versus running out of cash in two. The "
        "Consortium Pivot is how Velos buys a third option — anchoring the standard that will "
        "eventually disrupt the freight market, instead of being disrupted by it.",
        "Happy to take questions.”",
    ]
    for p in rec:
        story.append(Paragraph(p, st["quote"]))

    story.append(Paragraph("Q&amp;A preparation", st["section"]))
    story.append(Paragraph(
        "Presenter B fields strategy/finance questions first; Presenter A handles framework/tech.",
        st["body"]
    ))
    qa = [
        ("Why would competitors join a Velos-led consortium?",
         "Because the alternative is each of them paying their own €15M for a proprietary standard "
         "no one adopts. Shared cost plus neutral governance is rational once one party offers it "
         "credibly. Banks and customs also need a single counterparty — they will pressure logistics "
         "firms toward consolidation."),
        ("Why 18 months and not 12 or 24?",
         "12 is too short to recruit governance partners through legal review. 24 risks burning Trust&#8217;s "
         "runway before milestones are testable. 18 maps to typical B2B consortium formation cycles "
         "and preserves enough Prime cash for a controlled spin-out if needed."),
        ("What if no consortium members sign?",
         "That is the gate. If the month-12 milestone misses, Velos initiates structured spin-out of "
         "Trust — strategic acquirer or PE — and harvests Prime. Optionality is preserved because "
         "Velos controls the timing."),
        ("Doesn&#8217;t this dilute Velos&#8217;s competitive advantage?",
         "Velos has no defensible advantage in proprietary Trust today — adoption is stalled, that is "
         "the whole problem. Opening the protocol trades a non-existent moat for two real ones: "
         "standard curator status and consortium founder leverage."),
        ("Why couldn&#8217;t Velos just push through to Model 3.0 directly with the consortium?",
         "Model 3.0 is a destination, not a starting move. Industry protocols emerge from working "
         "ecosystems — AWS came after Amazon&#8217;s e-commerce platform proved at scale; Ethereum "
         "came after Bitcoin demonstrated that distributed consensus actually works. Velos must "
         "first establish that the consortium can run cross-firm transactions reliably — that is "
         "Model 2.0. Only then does &#8216;Velos sets the rules&#8217; carry weight. Skipping Model 2.0 "
         "is exactly what got them stuck in the first place."),
        ("How do you defend against larger players (Maersk, DHL) hijacking your consortium?",
         "First-mover governance leverage. Velos sets founding rules, equity splits, and IP terms. "
         "A larger entrant joins on Velos&#8217;s terms or builds their own — and a second standard "
         "fragments the market against them. Velos becomes too entangled to dislodge."),
    ]
    for q, a in qa:
        story.append(Paragraph(f"Q. {q}", st["qa_q"]))
        story.append(Paragraph(f"A. {a}", st["qa_a"]))

    story.append(Spacer(1, 0.6 * cm))
    story.append(Paragraph(
        "<i>Pacing target: Slide 1 ≈ 0:30 · Slide 2 ≈ 4:00 · Slide 3 ≈ 4:00 · Q&amp;A buffer ≈ 1:30 · "
        "Total ≈ 10:00.</i>",
        st["footnote"]
    ))

    doc.build(story)
    print(f"  Saved: {path}")


# =============================================================================
# MAIN
# =============================================================================

if __name__ == "__main__":
    out = os.path.dirname(os.path.abspath(__file__))
    print("Building Velos deliverables...")
    build_full_deck(os.path.join(out, "Velos_Full_Deck.pptx"))
    build_presentation_deck(os.path.join(out, "Velos_Presentation.pptx"))
    build_speaker_pdf(os.path.join(out, "Velos_Speaker_Scripts.pdf"))
    print("Done.")
