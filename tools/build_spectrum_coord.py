#!/usr/bin/env python3
"""Last workbook sheet: LTE Spectrum Coordination (CloudAIR channel cloudification).

The dedicated Huawei FPD is referenced as CA FPD Ch.23 item 55 but is not in this
repository. This sheet is built only from:
  - CA FPD eRAN21.1 Issue 11 (function-impact rows + reference list)
  - CaMgtCfg.CellCaAlgoSwitch option meaning published in related eRAN FPDs
  - LTE FDD and NR Spectrum Sharing FPD (DSS impact)
  - EPC-based NSA FPD (reserved bit + enhancement vs NSA PCC anchoring)
  - Huawei public CloudAIR / Turkcell commercial-trial description

Undocumented MML samples, counter IDs, and license models are NOT invented.
"""
import os, sys
sys.path.insert(0, os.path.dirname(__file__))
from ca_excel_style import *
from build_step1_2cc import (
    add_param_header, add_param, add_mml_header, add_mml, add_ctr_header, add_ctr,
    FIG, COLS, _fig,
)
from make_spectrum_coord_figures import make_ul_dl_decoupling, make_pcell_flow


def build_spectrum_coord(wb):
    os.makedirs(FIG, exist_ok=True)
    fig_ul = make_ul_dl_decoupling()
    fig_flow = make_pcell_flow()

    ws = wb.create_sheet("7. LTE Spectrum Coordination")
    setup_sheet(ws, "Spectrum Coord")
    set_widths(ws, [8, 28, 34, 24, 30, 56, 44, 12, 12, 12])
    r = 1
    r = banner(ws, r, COLS, "  7.  LTE Spectrum Coordination   —  UL/DL decoupling on a CA UE  (CloudAIR)")
    r = note_bar(
        ws, r, COLS,
        "Place in the CA FPD: Chapter 23 Reference Documents, item 55. “LTE Spectrum Coordination”.  "
        "The dedicated Feature Parameter Description is a separate eRAN document and is NOT attached to this repository.  "
        "This sheet therefore uses only statements that appear in the CA FPD, in related Huawei FPDs (DSS, NSA, CellCaAlgoSwitch), "
        "and in Huawei public CloudAIR / Turkcell trial material.  "
        "Rows marked “Reconstructed MML” use the same MOD CAMGTCFG bit-switch syntax as the CA FPD samples; they are not copied from a Spectrum Coordination MML example.  "
        "License models and dedicated counter IDs are not published in the sources below — see License Control Item Lists and the dedicated FPD."
    )
    ws.row_dimensions[r - 1].height = 72

    # ------------------------------------------------------------------ A
    r = section(ws, r, COLS, "A.  Principal  —  what the feature is and how it sits on Carrier Aggregation")
    r = subsection(ws, r, COLS, "A1.  One-sentence definition")
    r = body(
        ws, r, COLS,
        "LTE Spectrum Coordination is Huawei CloudAIR “channel cloudification” for LTE: it lifts the binding between "
        "the uplink channel and the downlink channel of a CA UE so that DL can stay on both the high band and the low band "
        "while UL is placed on the band that currently gives the better uplink experience.  "
        "Huawei’s published parameter meaning: when SpectrumCoordinationSwitch is selected, PCell changes are triggered based on uplink quality.  "
        "Applies to LTE FDD and LTE TDD.  Default of the switch is Off."
    )
    r = subsection(ws, r, COLS, "A2.  Why it exists (the coverage mismatch CA does not solve by itself)")
    r = bullets(ws, r, COLS, [
        "High bands (1800 / 2100 / TDD mid-band) have more spectrum and higher DL peak, but worse UL coverage: UE Tx power is limited (~23 dBm) and path loss + penetration rise with frequency.",
        "Low bands (700 / 800 / 900) have excellent UL coverage but less bandwidth. A UE that camps on the high band as PCell keeps PUCCH/PUSCH on that high-band PCell (unless UL CA is also configured).",
        "Ordinary DL CA already gives the UE both carriers for downlink. It does not by itself move the UL control/data to the low band when the UE is at the cell edge — PCell is still chosen mainly by DL measurements / PCC anchoring.",
        "Spectrum Coordination adds UL-quality-triggered PCell change on top of CA, so the cell-edge UE can keep DL aggregation and still uplink on the coverage layer.",
        "Huawei commercial trial (Turkcell Antalya, eRAN 13.1, Jan 2018): DL hosted by both bands; cell-center UE selects high band as UL; cell-edge UE selects low band as UL. Reported cell-edge downlink data rate increase > 30%.",
        "CloudAIR white paper: usable for FDD high+low, and also for TDD high band paired with FDD low band.",
    ])
    ws.row_dimensions[r - 1].height = 130
    r = insert_figure(ws, r, fig_ul, COLS,
                      "Figure SC-1  UL/DL decoupling on a CA UE  (schematic reconstructed from Huawei CloudAIR / Turkcell description — not a Huawei FPD original)")

    r = subsection(ws, r, COLS, "A3.  Relationship to Carrier Aggregation (this workbook)")
    r = headers(ws, r, ["Layer", "What CA already does (sheets 1–6)", "What Spectrum Coordination adds"] + [""] * 7)
    rel = [
        ("DL user-plane", "Aggregates 2–5 (up to 8) CCs after SCell activation (MAC CE). Peak = sum of CCs, capped by BBP and UE category.",
         "Does not replace CA. DL remains hosted by both high and low bands so cell-edge DL still uses the high-band SCell."),
        ("UL user-plane / control", "Default: UL on PCell/PCC. UL 2CC is a separate feature (CA FPD Ch.10, CaUl2CCSwitch).",
         "Chooses which cell is PCell according to UL quality, so UL rides the better uplink band without requiring UL CA."),
        ("PCell selection", "PCC anchoring (A1 then A5) + preferred PCell priority / PccFreqCfg. Driven by DL RSRP/RSRQ.",
         "Additional PCell-change trigger: uplink quality. Can move PCell from high band (center) to low band (edge) and back."),
        ("SCell management", "Ch.4.6 A4 add / A6 change / A2 remove / traffic+CQI activate/deactivate. Unchanged as the state machine.",
         "After a UL-quality PCell change, the previous PCell typically remains as SCell so DL CA continues. Do not reuse A1–A6 as the Spectrum Coordination trigger."),
        ("Licensing", "LAOFD / TDLAOFD / LEOFD CA licenses in Step1–4.",
         "Spectrum Coordination is a separate eRAN feature (FDD and TDD names appear on CellCaAlgoSwitch). CA licenses still required because DL CA is the vehicle."),
    ]
    for i, rec in enumerate(rel):
        vals = list(rec) + [""] * 7
        r = table_row(ws, r, vals, fills=[alt_fill(i)] * 10, height=48)
        merge(ws, r - 1, 3, r - 1, COLS)

    r = subsection(ws, r, COLS, "A4.  What this feature is NOT  (do not mix these on the same change request)")
    r = headers(ws, r, ["Look-alike name", "What it actually is", "Typical switch", "This sheet?"] + [""] * 6)
    nots = [
        ("LTE Spectrum Coordination (this sheet)", "UL/DL decoupling across LTE high/low bands on a CA UE; PCell change on UL quality.",
         "CaMgtCfg.CellCaAlgoSwitch → SpectrumCoordinationSwitch", "YES"),
        ("LTE spectrum coordination enhancement", "WBB / RRN enhancement of the above. NSA PCC anchoring takes precedence if both on.",
         "CaMgtCfg.CellCaAlgoSwitch → WbbCaMultiCarrierCoordSw", "YES (section D/E)"),
        ("LTE FDD and NR Flash DSS / Hybrid DSS / SUL DSS", "4G and 5G share the SAME carrier’s RBs over time. Reduces LTE UL RBs → fewer UEs get Spectrum Coordination.",
         "SpectrumCloud.SpectrumCloudSwitch = LTE_NR_SPECTRUM_SHR (or SUL variant)", "Impact only"),
        ("GSM and LTE Spectrum Concurrency / UMTS–LTE Spectrum Sharing", "CloudAIR spectrum cloudification between RATs on one band. CA FPD: do not use 5 MHz / GL-concurrency cells as PCell (PUCCH vs SRS).",
         "SpectrumCloudSwitch = GL_SPECTRUM_CONCURRENCY or UL_SPECTRUM_SHARING", "Impact only"),
        ("2019 “LTE&NR Spectrum Coordination” press note", "Dynamic 4G/5G spectrum allocation on C-band / 2.6 GHz and FDD 2.1 GHz sharing. Different product story from LTE high/low UL decoupling.",
         "LTE–NR DSS / sharing family — not SpectrumCoordinationSwitch", "NO — related but different"),
        ("5G UL/DL decoupling (NR SUL / NR FPD item 75 in CA refs)", "NR uplink on a supplementary or low band while NR DL stays on C-band. Different RAT, different FPD.",
         "NR feature documentation “UL and DL Decoupling in 5G RAN”", "NO"),
    ]
    for i, rec in enumerate(nots):
        vals = list(rec) + [""] * 6
        r = table_row(ws, r, vals, fills=[PALE_GREEN if i == 0 else alt_fill(i)] * 10, height=44)
        merge(ws, r - 1, 4, r - 1, COLS)

    # ------------------------------------------------------------------ B
    r = section(ws, r, COLS, "B.  Triggering conditions  (UL-quality PCell change — not CA A1–A6)")
    r = note_bar(
        ws, r, COLS,
        "The dedicated FPD’s UL-quality threshold table is not in this repository. What is documented: "
        "“Spectrum coordination is enabled only if this option is selected. With spectrum coordination enabled, PCell changes are triggered based on uplink quality. This option applies only to LTE FDD and LTE TDD.”  "
        "Default: Off.  Worked numbers below are radio-propagation illustrations so the operations team can see the size of the UL gap; they are not Huawei factory defaults."
    )
    r = insert_figure(ws, r, fig_flow, COLS,
                      "Figure SC-2  Enablement and UL-quality PCell change  (reconstructed from CellCaAlgoSwitch meaning + CloudAIR; not a Huawei original figure)")

    r = subsection(ws, r, COLS, "B1.  Preconditions before the UL-quality trigger can fire")
    r = bullets(ws, r, COLS, [
        "Downlink CA is already working on the high/low pair (this workbook Step1 as a minimum). Spectrum Coordination does not aggregate carriers by itself.",
        "The two cells overlap in coverage. Neighbors exist (CA FPD: NoRmvFlag = FORBID_RMV_ENUM on CA neighbors).",
        "UE is CA-capable (Rel-10+ for 2CC). Same UE-attribute exclusions as CA still apply (emergency call, eMBMS, eMTC, high-speed rules — see Step1 B1).",
        "SpectrumCoordinationSwitch = On on the cells that will participate (PCell and the partner band). Default is Off, so nothing happens until this bit is selected.",
        "Enough LTE uplink RBs remain on the cell. DSS and UL multi-cluster each reduce the LTE UL RB pool and therefore reduce the proportion of UEs for which Spectrum Coordination takes effect (DSS FPD).",
        "For NSA DC UEs: if RsvdSwPara6_bit20 is selected after Spectrum Coordination is on, the feature is disabled for those UEs (NSA FPD) — they will not be HO’d from PCC to a non-candidate PCC.",
    ])
    ws.row_dimensions[r - 1].height = 120

    r = formula_box(
        ws, r, COLS,
        "B2.  Size of the UL problem  —  extra path loss of the high band vs the low band",
        "ΔPL_freq (dB) ≈ 20 · log10(f_high / f_low)     (same distance, free-space / Okumura-style frequency term)\n"
        "ΔPL_total ≈ ΔPL_freq + ΔPL_penetration + ΔPL_clutter\n"
        "UL SNR_high ≈ (P_UE − PL_high − N) ;   UL SNR_low ≈ (P_UE − PL_low − N)\n"
        "With P_UE capped (typically 23 dBm), the high band runs out of UL SNR first. That is the cell-edge trigger region.",
        "Sample pair used on many LTE CA networks (not a Huawei MML value):\n"
        "  High UL ≈ Band 3  1747.5 MHz     Low UL ≈ Band 8  897.5 MHz\n"
        "  ΔPL_freq = 20 · log10(1747.5 / 897.5) = 20 · log10(1.947) ≈ 5.8 dB\n"
        "  Extra building penetration 1800 vs 900 is typically another ~5 to 10 dB  →  ΔPL_total ≈ 11 to 16 dB at the same location.\n"
        "  If a cell-edge UE has UL SINR = −2 dB on the high-band PCell, the same UE on the low band is in the order of +9 to +14 dB, all else equal.",
        "Spectrum Coordination’s job is to change PCell to the low-band cell in this region so PUCCH/PUSCH ride the coverage layer, "
        "while the high-band cell stays as SCell and continues to carry DL. When the UE walks back to the center and high-band UL recovers, PCell can return to the high band."
    )

    r = subsection(ws, r, COLS, "B3.  Worked examples  (operator illustration — not Huawei factory thresholds)")
    r = headers(ws, r, ["SN", "UE location", "High-band UL SINR (sample)", "Low-band UL SINR (sample)", "Expected PCell after coordination", "DL CA state", "Notes"] + [""] * 3)
    ex = [
        ("1", "Cell center", "+18 dB (good)", "+22 dB (also good)", "HIGH band (capacity layer)",
         "Both CCs active", "UL does not need the coverage layer. Keep PCell on high band so UL uses the typically wider high-band PUSCH."),
        ("2", "Mid cell", "+6 dB", "+14 dB", "Depends on live UL-quality criterion in the dedicated FPD",
         "Both CCs active", "Do not invent a dB threshold. Confirm hysteresis / time-to-trigger in LTE Spectrum Coordination FPD / Parameter Reference to avoid PCell ping-pong."),
        ("3", "Cell edge (Turkcell-class)", "−2 dB (UL limited)", "+10 dB (usable)", "LOW band (coverage layer)",
         "Both CCs active — high band remains SCell for DL",
         "This is the gain region: cell-edge DL rate rose >30% in the 2018 commercial trial because DL still used the high-band SCell while UL survived on low band."),
        ("4", "Cell edge, CA not configured", "−2 dB", "+10 dB", "Idle/connected reselection or coverage HO only — no SC",
         "Single carrier", "Without CA, moving PCell to low band also loses high-band DL. Spectrum Coordination is not a substitute for 2CC."),
        ("5", "NSA DC UE, bit20 = 1", "any", "any", "PCC not moved by Spectrum Coordination",
         "LTE CA + NR SCG as NSA allows",
         "NSA FPD: RsvdSwPara6_bit20 disables spectrum coordination for NSA DC UEs to prevent HO from PCC to a non-candidate PCC."),
        ("6", "NSA DC PCC anchoring ON + WbbCaMultiCarrierCoordSw ON", "any", "any", "NSA PCC anchoring wins",
         "NSA DC", "NSA FPD: NSA DC PCC anchoring takes precedence over LTE spectrum coordination enhancement. Only NSA anchoring takes effect."),
        ("7", "DSS cell, few LTE UL RBs left", "measured", "measured", "Feature may not take effect for this UE",
         "CA may still run on remaining LTE RBs",
         "DSS FPD: fewer uplink RBs available for LTE reduces the proportion of UEs for which LTE Spectrum Coordination takes effect."),
    ]
    for i, rec in enumerate(ex):
        vals = list(rec) + [""] * 3
        r = table_row(ws, r, vals, fills=[EXAMPLE if i in (0, 2) else alt_fill(i)] * 10, height=52)
        merge(ws, r - 1, 7, r - 1, COLS)

    r = subsection(ws, r, COLS, "B4.  After the PCell change  —  CA state machine continues")
    r = body(
        ws, r, COLS,
        "Once the low-band cell is PCell, ordinary CA Ch.4.6 still applies: the high-band neighbor should already be a candidate SCell "
        "(group CaGroupSCellCfg or adaptive SccFreqCfg), A4/blind add can keep it configured, and MAC CE activation keeps DL on both CCs.  "
        "Event A5 Thresh1 remains −43 dBm / −3 dB for CA PCC anchoring; A6 Hys remains 1 dB. Those constants are CA FPD facts and are not Spectrum Coordination thresholds.  "
        "If the operations team also has Enhanced PCC anchoring ON, both DL-based anchoring and UL-quality PCell change can request a PCell change. "
        "The dedicated Spectrum Coordination FPD is the place that states the exact precedence; this repository does not. "
        "Known documented precedence: NSA DC PCC anchoring > LTE spectrum coordination enhancement."
    )

    # ------------------------------------------------------------------ C
    r = section(ws, r, COLS, "C.  Leaving / rollback")
    r = headers(ws, r, ["SN", "Leave path", "What happens", "How to do it / what to watch"] + [""] * 6)
    leave = [
        ("1", "UE moves back to cell center", "High-band UL quality recovers → PCell can return to high band; low band stays as SCell for DL.",
         "Watch PCell HO counters (L.HHO.*.CAUser.PCC). Confirm no ping-pong (HO prep ≫ HO success). Dedicated hysteresis is in the Spectrum Coordination FPD."),
        ("2", "SCell (high band) radio quality collapses", "Ordinary CA A2 / traffic / CQI deactivation / RRC remove. UE may keep low-band PCell as a single carrier.",
         "Same leaving package as Step1 section C. Not Spectrum Coordination-specific."),
        ("3", "Operator turns the feature off", "UL-quality PCell changes stop. Existing CA continues with DL-based PCC anchoring / A4–A6 only.",
         "MOD CAMGTCFG CellCaAlgoSwitch=SpectrumCoordinationSwitch-0 on every cell where it was enabled. Then re-check PCell distribution vs band."),
        ("4", "NSA DC UE, bit20 selected", "Spectrum Coordination does not move that UE’s PCC.",
         "MOD ENBCELLRSVDPARA RsvdSwPara6_bit20-1. Confirm exact bit MML in eNodeBFunction Used Reserved Parameter List."),
        ("5", "DSS / UL multi-cluster shrinks LTE UL RBs", "Feature takes effect for a smaller fraction of UEs — looks like “leaving” in KPI even though the switch is still On.",
         "Check remaining LTE UL RB count and DSS sharing ratio before blaming Spectrum Coordination."),
        ("6", "Load-balance inter-CC transfer vs enhancement", "CA FPD: inter-CC load transfer does not take effect for WBB CA UEs or RRNs on which LTE spectrum coordination enhancement has taken effect.",
         "If you need DlCaLbAlgoSwitch / CaLoadBalancePreAllocSwitch on those WBB/RRN UEs, do not expect both to win."),
    ]
    for i, rec in enumerate(leave):
        vals = list(rec) + [""] * 6
        r = table_row(ws, r, vals, fills=[alt_fill(i)] * 10, height=48)
        merge(ws, r - 1, 4, r - 1, COLS)

    # ------------------------------------------------------------------ D
    r = section(ws, r, COLS, "D.  Prerequisite / exclusive / impacted features  (one feature per row; related parameters grouped)")
    r = note_bar(ws, r, COLS, "CA FPD function-impact tables (Ch.5 / Ch.13 CloudAIR rows) plus DSS FPD and NSA FPD. Related parameters of the same feature are grouped in the Parameter column.")
    r = headers(ws, r, ["SN", "Relation", "RAT", "Feature", "Related parameter(s)", "Impact / rule", "Notes"] + [""] * 3)
    feats = [
        ("1", "Prerequisite", "FDD/TDD", "Downlink 2CC CA (this workbook Step1)",
         "CaMgtCfg.CellCaAlgoSwitch (2CC bits); ENodeBAlgoSwitch.CaAlgoSwitch FreqCfgSwitch/AdpCaSwitch; CaGroup or PccFreqCfg+SccFreqCfg",
         "DL must already be aggregatable on the high+low pair. Spectrum Coordination does not create CA.",
         "3CC/4CC/5CC optional; they increase DL after the UL PCell is correct."),
        ("2", "Prerequisite", "FDD/TDD", "Overlapping high/low coverage + CA neighbor",
         "EutranInterNFreq.AggregationAttribute FREQ_MEAS_FLAG; EutranCellMeasurement.NoRmvFlag=FORBID_RMV_ENUM",
         "Without overlap, a PCell change to the other band drops the UE out of coverage of the original DL CC.",
         "Same coverage rule as CA Ch.5.3.5."),
        ("3", "This function", "FDD/TDD", "LTE Spectrum Coordination",
         "CaMgtCfg.CellCaAlgoSwitch → SpectrumCoordinationSwitch",
         "Master enable. Default Off. PCell changes based on uplink quality when selected.",
         "CellCaAlgoSwitch description in related eRAN FPDs. Dedicated FPD = CA ref 55."),
        ("4", "Enhancement", "TDD (documented in CA FPD CloudAIR rows); listed generally on CellCaAlgoSwitch",
         "LTE spectrum coordination enhancement",
         "CaMgtCfg.CellCaAlgoSwitch → WbbCaMultiCarrierCoordSw",
         "CA FPD Ch.5 / Ch.13: inter-CC load transfer triggered by cell load (DlCaLbAlgoSwitch) or by cell-load differences (CaLoadBalancePreAllocSwitch) does NOT take effect for WBB CA UEs or RRNs on which this enhancement has taken effect.",
         "NSA FPD: NSA DC PCC anchoring takes precedence over this enhancement."),
        ("5", "Impacted by", "FDD", "LTE FDD and NR Flash / Hybrid / SUL Dynamic Spectrum Sharing",
         "SpectrumCloud.SpectrumCloudSwitch = LTE_NR_SPECTRUM_SHR or LTE_NR_UPLINK_SPECTRUM_SHR; CaMgtCfg.CellCaAlgoSwitch → SpectrumCoordinationSwitch",
         "DSS FPD: the number of uplink RBs available for LTE decreases, which reduces the proportion of UEs for which LTE Spectrum Coordination takes effect.",
         "CA FPD also: DSS cells not recommended as PCell (PUCCH so large SRS cannot be configured)."),
        ("6", "Impacted by", "FDD/TDD", "UL Multi-Cluster Scheduling",
         "CellAlgoSwitch.UlSchExtSwitch → UlMultiClusterSwitch",
         "DSS FPD lists this next to Spectrum Coordination: fewer LTE UL RBs → fewer UEs get Spectrum Coordination, and UL frequency-selective / multi-cluster gain also drops.",
         "Review this bit on DSS-sharing cells."),
        ("7", "Impacted by / exception", "FDD/TDD", "EPC-based NSA DC",
         "ENBCELLRSVDPARA.RsvdSwPara6 → RsvdSwPara6_bit20; CaMgtCfg.CellCaAlgoSwitch → WbbCaMultiCarrierCoordSw; NSA PCC-anchoring switches (see NSA FPD)",
         "After Spectrum Coordination is enabled, select bit20 to disable it for NSA DC UEs (prevents HO from PCC to non-candidate PCC). NSA PCC anchoring > coordination enhancement.",
         "Reserved-parameter details: eNodeBFunction Used Reserved Parameter List."),
        ("8", "Impact (does not take effect together)", "TDD WBB/RRN", "Inter-CC load transfer",
         "ENodeBAlgoSwitch.CaLbAlgoSwitch → DlCaLbAlgoSwitch; ENodeBAlgoSwitch.CaAlgoSwitch → CaLoadBalancePreAllocSwitch; WbbCaMultiCarrierCoordSw",
         "Load-triggered inter-CC transfer does not take effect on WBB CA UEs / RRNs where spectrum coordination enhancement has taken effect.",
         "CA FPD CloudAIR impact rows, DL 2CC and FDD+TDD CA chapters."),
        ("9", "Related CloudAIR (different feature)", "FDD", "GSM and LTE Spectrum Concurrency",
         "SpectrumCloud.SpectrumCloudSwitch = GL_SPECTRUM_CONCURRENCY",
         "CA FPD: such LTE cells are not recommended as PCells (PUCCH vs SRS). Relaxed-backhaul: if they are SCells, pre-scheduling RB estimate can be stale.",
         "Do not confuse with SpectrumCoordinationSwitch."),
        ("10", "Related CloudAIR (different feature)", "FDD", "UMTS and LTE Spectrum Sharing (and DC-HSDPA variant)",
         "SpectrumCloud.SpectrumCloudSwitch = UL_SPECTRUM_SHARING or DC_HSDPA_BASED_UL_SPECTRUM_SHR",
         "CA FPD: 5 MHz cells not recommended as PCell. Massive CA not recommended on sharing cells.",
         "Spectrum cloudification, not channel cloudification."),
        ("11", "Related CA (still required)", "FDD/TDD", "Enhanced PCC anchoring / A5 HO events",
         "ENodeBAlgoSwitch.CaAlgoSwitch → EnhancedPccAnchorSwitch; CaAlgoExtSwitch → CaA5HoEventSwitch + CaA5HoEventEnhSwitch",
         "DL-based PCell steering remains. Spectrum Coordination adds UL-quality PCell change on top. Keep A5 HO switches so a UE can HO from PCell to its SCell (CA Ch.4.3 recommendation).",
         "Do not turn CA PCC anchoring off unless the dedicated FPD says so — it is not stated here."),
        ("12", "Hardware / product", "FDD/TDD", "3900 / 5900 series eNodeB (same family as CA)",
         "—",
         "CA FPD hardware chapters apply to the CA vehicle. Dedicated Spectrum Coordination hardware notes are in its own FPD (not in repo).",
         "Contact Huawei service for board combination if activating on LampSite / RRU-only sites."),
    ]
    for rec in feats:
        r = table_row(ws, r, list(rec) + [""] * 3,
                      fills=[PALE_GREEN if rec[1] == "Prerequisite" else PALE_GOLD if rec[1].startswith("This") else PALE_ORANGE if rec[1].startswith("Impact") else alt_fill(int(rec[0]))] * 10,
                      bolds=[False, True, False, True, False, False, False],
                      height=62)
        merge(ws, r - 1, 7, r - 1, COLS)

    # ------------------------------------------------------------------ E
    r = section(ws, r, COLS, "E.  Parameters  (SN, MO, name, default, recommended, description, notes)")
    r = note_bar(ws, r, COLS, "Only parameters that appear in the sources listed in section L. Factory defaults that the source actually prints are used; otherwise “See Parameter Reference”. Recommended values are operational, not Huawei-confidential tuning.")
    r = add_param_header(ws, r)
    r = add_param(ws, r, 1, "CaMgtCfg", "CellCaAlgoSwitch / SpectrumCoordinationSwitch", "Off", "1 on every high-band and low-band cell that should participate, after Step1 CA is verified",
                  "Spectrum coordination is enabled only if this option is selected. With spectrum coordination enabled, PCell changes are triggered based on uplink quality. Applies to LTE FDD and LTE TDD.",
                  "Master switch. Same MO as CaDl3CCSwitch / CaUl2CCSwitch / … Default Off is printed in related FPD CellCaAlgoSwitch listing.")
    r = add_param(ws, r, 2, "CaMgtCfg", "CellCaAlgoSwitch / WbbCaMultiCarrierCoordSw", "See Parameter Reference", "1 only if WBB/RRN coordination enhancement is required AND NSA PCC anchoring is not supposed to win",
                  "LTE spectrum coordination enhancement. Referenced from CA FPD CloudAIR impact rows and from NSA FPD (“coordination enhancement”).",
                  "Related: DlCaLbAlgoSwitch, CaLoadBalancePreAllocSwitch (load transfer does not take effect on WBB CA UEs/RRNs once enhancement has taken effect). Related: NSA PCC anchoring takes precedence.", related=True)
    r = add_param(ws, r, 3, "ENBCELLRSVDPARA", "RsvdSwPara6 / RsvdSwPara6_bit20", "See reserved-parameter list (typically 0 = not selected)", "1 on NSA sites if NSA DC UEs must NOT be moved by Spectrum Coordination",
                  "After LTE spectrum coordination is enabled, this bit can be selected to disable spectrum coordination for NSA DC UEs, i.e. to prevent NSA DC UEs from being handed over from PCCs to non-candidate PCCs.",
                  "Source: EPC-based NSA FPD. Confirm meaning, version, and MML in 3900 & 5900 eNodeBFunction Used Reserved Parameter List. Related: SpectrumCoordinationSwitch must already be On.", related=True)
    r = add_param(ws, r, 4, "CellAlgoSwitch", "UlSchExtSwitch / UlMultiClusterSwitch", "See Parameter Reference", "Review on DSS / tight-UL cells; do not assume Off",
                  "UL multi-cluster scheduling. DSS FPD: fewer LTE UL RBs reduces the proportion of UEs for which Spectrum Coordination takes effect, and reduces UL multi-cluster / frequency-selective gain.",
                  "Related to SpectrumCoordinationSwitch only through the shared UL RB pool — not an alias of it.", related=True)
    r = add_param(ws, r, 5, "SpectrumCloud", "SpectrumCloudSwitch", "See Parameter Reference (DSS Off unless DSS is the job)", "Keep DSS and Spectrum Coordination planning coupled if both are on the same LTE cell",
                  "LTE_NR_SPECTRUM_SHR / LTE_NR_UPLINK_SPECTRUM_SHR / GL_SPECTRUM_CONCURRENCY / UL_SPECTRUM_SHARING are different CloudAIR functions. Flash DSS reduces LTE UL RBs (Spectrum Coordination UE proportion ↓). GL concurrency / UMTS sharing: CA FPD says do not use those cells as PCell if PUCCH kills SRS.",
                  "Do not set SpectrumCloudSwitch to enable LTE Spectrum Coordination — that is CellCaAlgoSwitch.", related=True)
    r = add_param(ws, r, 6, "ENodeBAlgoSwitch", "CaLbAlgoSwitch / DlCaLbAlgoSwitch", "See Parameter Reference", "Know the conflict on WBB/RRN if WbbCaMultiCarrierCoordSw is On",
                  "Inter-CC load transfer triggered by cell load. CA FPD: does not take effect for WBB CA UEs or RRNs on which LTE spectrum coordination enhancement has taken effect.",
                  "Related: WbbCaMultiCarrierCoordSw.", related=True)
    r = add_param(ws, r, 7, "ENodeBAlgoSwitch", "CaAlgoSwitch / CaLoadBalancePreAllocSwitch", "See Parameter Reference", "Same awareness as DlCaLbAlgoSwitch",
                  "Inter-CC load transfer triggered by cell-load differences. Same CA FPD exclusion versus spectrum coordination enhancement on WBB CA UEs / RRNs.",
                  "Related: WbbCaMultiCarrierCoordSw, DlCaLbAlgoSwitch.", related=True)
    r = add_param(ws, r, 8, "ENodeBAlgoSwitch", "CaAlgoSwitch / EnhancedPccAnchorSwitch", "See Parameter Reference (CA Step1 recommends select)", "Keep as in Step1 unless dedicated FPD says otherwise",
                  "DL-measurement PCC anchoring (A1/A5). Still the CA way of choosing a better PCell. Spectrum Coordination adds a UL-quality reason to change PCell.",
                  "Related CA: PccAnchorSwitch (ignored if Enhanced is On), PccFreqCfg.PccA4RsrpThd, CaGroup PreferredPCellPriority.", related=True)
    r = add_param(ws, r, 9, "ENodeBAlgoSwitch", "CaAlgoExtSwitch / CaA5HoEventSwitch + CaA5HoEventEnhSwitch", "See Parameter Reference", "Select BOTH (CA FPD Ch.4.3 recommendation)",
                  "Changes HO event A4 to A5 when SCells exist so the UE can be handed over from its PCell to its SCell. Spectrum Coordination’s PCell change between high and low bands needs this HO path to be legal.",
                  "Related: InterFreqHoGroup A5 thresholds. A5 Thresh1 is always −43 dBm / −3 dB (CA FPD).", related=True)
    r = add_param(ws, r, 10, "CaMgtCfg", "CellCaAlgoSwitch / CaUl2CCSwitch", "Off", "Independent — enable only if UL 2CC is planned",
                  "Uplink 2CC aggregation (CA FPD Ch.10). Spectrum Coordination is an alternative/complement for UL coverage, not a substitute for UL 2CC peak rate.",
                  "If both are on, the dedicated FPDs must be read together. Not described further here.", related=True)
    r = add_param(ws, r, 11, "EutranInterNFreq", "AggregationAttribute / FREQ_MEAS_FLAG", "Site-planned", "Selected on the partner-band frequency",
                  "Frequency is a CA candidate. If deselected, the other band cannot be SCell after a PCell change.",
                  "CA Step1 prerequisite, still required here.", related=True)
    r = add_param(ws, r, 12, "EutranCellMeasurement", "NoRmvFlag", "See Parameter Reference", "FORBID_RMV_ENUM on CA neighbors of the pair",
                  "Stops ANR from deleting the high/low neighbor used for CA and for PCell swap.",
                  "CA FPD neighbor rule.", related=True)

    r = subsection(ws, r, COLS, "E2.  Parameters that exist only in the dedicated FPD  —  not printed here")
    r = bullets(ws, r, COLS, [
        "UL-quality metric, entering/leaving thresholds, hysteresis, time-to-trigger, and any blacklist (VoLTE / high-speed / QCI) for Spectrum Coordination PCell change.",
        "Whether the feature requires group-based CA, adaptive CA, or both.",
        "Dedicated performance counters and license models (Feature ID / Model / sales unit).",
        "Action: open eRAN Feature Documentation → LTE Spectrum Coordination, then filter eNodeBFunction Parameter Reference by that feature ID (CA FPD Ch.20 FAQ).",
    ], fill_hex=PALE_RED)
    ws.row_dimensions[r - 1].height = 72

    # ------------------------------------------------------------------ F
    r = section(ws, r, COLS, "F.  MML  (reconstructed from documented switches; not a dedicated-FPD sample)")
    r = note_bar(ws, r, COLS, "Bit-switch syntax matches CA FPD samples (CellCaAlgoSwitch=Name-1). Replace LocalCellId and confirm reserved-parameter MML on the live version. Doc example eNodeBId=1234 is not required here.")
    r = add_mml_header(ws, r)
    mmls = [
        (1, "0", "Both", "FDD/TDD", "Pre-check",
         "DSP CELLCA; DSP CAMGTCFG: LocalCellId=0; LST CAMGTCFG: LocalCellId=0;",
         "Confirm Step1 CA is already producing SCell add/act counters before enabling Spectrum Coordination."),
        (2, "1", "Both", "FDD/TDD", "Enable SC",
         "MOD CAMGTCFG: LocalCellId=0, CellCaAlgoSwitch=SpectrumCoordinationSwitch-1;",
         "Reconstructed MML. Run on each participating high-band and low-band cell (repeat LocalCellId). Default was Off."),
        (3, "2", "Both", "FDD/TDD", "Enable SC",
         "MOD CAMGTCFG: LocalCellId=1, CellCaAlgoSwitch=SpectrumCoordinationSwitch-1;",
         "Partner-band cell. Use live LocalCellId (doc 2CC TDD examples used 0 and 1)."),
        (4, "3", "Both", "TDD WBB/RRN", "Optional enhancement",
         "MOD CAMGTCFG: LocalCellId=0, CellCaAlgoSwitch=WbbCaMultiCarrierCoordSw-1;",
         "Reconstructed. Only if enhancement is required. Conflicts with inter-CC load transfer on WBB CA UEs/RRNs. NSA PCC anchoring will override it."),
        (5, "4", "Both", "FDD/TDD NSA", "NSA exception",
         "MOD ENBCELLRSVDPARA: LocalCellId=0, RsvdSwPara6=RsvdSwPara6_bit20-1;",
         "Reconstructed from NSA FPD wording. Confirm bit name and command in Used Reserved Parameter List before live use."),
        (6, "5", "Both", "FDD/TDD", "UL RB awareness",
         "LST CELLALGOSWITCH: LocalCellId=0;",
         "Inspect UlSchExtSwitch / UlMultiClusterSwitch. Do not toggle blindly; DSS FPD only states the RB-pool impact."),
        (7, "6", "Both", "FDD", "DSS coexistence check",
         "LST SPECTRUMCLOUD: LocalCellId=0;",
         "If SpectrumCloudSwitch is LTE_NR_SPECTRUM_SHR (or SUL), expect a smaller Spectrum Coordination UE proportion."),
        (8, "7", "Both", "FDD/TDD", "HO path",
         "LST ENODEBALGOSWITCH:;",
         "Confirm CaA5HoEventSwitch and CaA5HoEventEnhSwitch remain selected (CA Ch.4.3) so PCell↔SCell HO is possible."),
        (9, "8", "Both", "FDD/TDD", "Verify",
         "DSP CAMGTCFG: LocalCellId=0;",
         "SpectrumCoordinationSwitch should show selected. Then watch CA PCell HO counters (section G)."),
        (10, "9", "Both", "FDD/TDD", "Rollback",
         "MOD CAMGTCFG: LocalCellId=0, CellCaAlgoSwitch=SpectrumCoordinationSwitch-0;",
         "Stops UL-quality PCell changes. Repeat on every cell. Enhancement bit and reserved bit20 are independent rollbacks."),
        (11, "10", "Both", "TDD WBB/RRN", "Rollback enhancement",
         "MOD CAMGTCFG: LocalCellId=0, CellCaAlgoSwitch=WbbCaMultiCarrierCoordSw-0;",
         "Restores the possibility of inter-CC load transfer on those WBB/RRN UEs (subject to DlCaLbAlgoSwitch / CaLoadBalancePreAllocSwitch)."),
        (12, "11", "Both", "FDD/TDD NSA", "Rollback NSA exception",
         "MOD ENBCELLRSVDPARA: LocalCellId=0, RsvdSwPara6=RsvdSwPara6_bit20-0;",
         "Re-enables Spectrum Coordination for NSA DC UEs if the master switch is still On. Confirm MML."),
    ]
    for rec in mmls:
        r = add_mml(ws, r, *rec)

    # ------------------------------------------------------------------ G
    r = section(ws, r, COLS, "G.  Counters  (CA FPD set used as the live monitor — dedicated SC counters not in repo)")
    r = note_bar(ws, r, COLS, "CA FPD Ch.21 says to filter eNodeBFunction Performance Counter Summary by the feature ID of LTE Spectrum Coordination. That ID is not printed in the CA FPD. Until the dedicated FPD is available, use the CA PCell HO / SCell / throughput counters below to see whether PCells are moving and whether cell-edge DL bits rise.")
    r = add_ctr_header(ws, r)
    ctrs = [
        (1, "1526728993", "L.HHO.ExecAttOut.CAUser.PCC", "CA UE PCell HO execution attempts", "Expect a rise when SC starts swapping high↔low PCell. Compare with non-SC cluster."),
        (2, "1526728994", "L.HHO.ExecSuccOut.CAUser.PCC", "CA UE PCell HO execution success", "Numerator of PCell HO success. Ping-pong if Att ≫ Succ or Att spikes with no KPI gain."),
        (3, "—", "L.HHO.PrepAttOut.CAUser.PCC (CA FPD HO family)", "CA UE PCell HO preparations", "Preparations without executions → measurement/threshold issue, not necessarily SC."),
        (4, "1526729045 / 9046", "L.CA.DLSCell.Add.Att / .Succ", "DL SCell add", "After PCell moves to low band, high band must still add as SCell. If add fails, you get coverage-layer UL but lose high-band DL."),
        (5, "1526728999 / 9000", "L.CA.DLSCell.Act.Att / .Succ", "SCell MAC activation", "DL decoupling only exists while the high-band SCell is active."),
        (6, "1526732658 / 1526729259", "L.CA.Traffic.bits.DL.PCell / .SCell", "DL bits on PCell vs SCell of CA UEs", "Cell-edge: more bits should appear on the high-band SCell while PCell is the low band."),
        (7, "1526729003 / 9004", "L.CA.DL.PCell.Act.Dur / SCell.Act.Dur", "Active duration", "Duration share high-as-SCell vs low-as-PCell is the operational fingerprint of SC."),
        (8, "—", "L.Thrp.bits.DL.CAUser / L.Thrp.Time.DL.RmvLastTTI.CAUser", "CA UE DL throughput ingredients", "Cell-edge bins (RSRP/SINR) should move if Turkcell-class gain is present."),
        (9, "1526729047 / 9048", "L.CA.DLSCell.Rmv.Att / .Succ", "SCell remove", "If SC PCell swap is followed by immediate SCell remove, overlap or A2 is too aggressive (Step1 leaving)."),
        (10, "—", "UL: PUSCH SINR / UL MCS / UL IBLER (cell performance / UE trace)", "UL quality of the current PCell", "This is the actual trigger family. Use MAE-Access UE trace on the same IMSI before/after the PCell swap."),
        (11, "DSP", "DSP UEONLINEINFO / RRC_CONN_RECFG", "sCellToAddModList after PCell band change; non-zero RB/TBS on both CCs", "Same verification style as CA Ch.5.4.2."),
        (12, "—", "Dedicated LTE Spectrum Coordination counters", "Not published in this repository", "Filter Performance Counter Summary by the feature ID from the dedicated FPD / License Control Item Lists."),
    ]
    for rec in ctrs:
        r = add_ctr(ws, r, *rec)

    # ------------------------------------------------------------------ H
    r = section(ws, r, COLS, "H.  KPI formulas")
    r = headers(ws, r, ["SN", "KPI", "Formula", "Unit", "Source / how to use"] + [""] * 5)
    kpis = [
        ("1", "CA UE PCell HO success rate",
         "L.HHO.ExecSuccOut.CAUser.PCC / L.HHO.ExecAttOut.CAUser.PCC × 100%",
         "%", "CA FPD Ch.5.4.3. Must stay healthy after SC is turned on; a drop means UL-quality PCell change is too aggressive or neighbors are missing."),
        ("2", "DL SCell add success (post-swap)",
         "L.CA.DLSCell.Add.Succ / L.CA.DLSCell.Add.Att × 100%",
         "%", "CA FPD. After PCell moves to low band this is the rate at which high-band DL is recovered as SCell."),
        ("3", "SCell activation success",
         "L.CA.DLSCell.Act.Succ / L.CA.DLSCell.Act.Att × 100%",
         "%", "CA FPD. Decoupling is real only while both CCs are active."),
        ("4", "CA UE average DL data rate (last-TTI removed)",
         "(L.Thrp.bits.DL.CAUser − L.Thrp.bits.DL.LastTTI.CAUser) / L.Thrp.Time.DL.RmvLastTTI.CAUser",
         "bit/s", "CA FPD. Split by RSRP band (cell-edge vs center) on the OSS. Cell-edge bin is the SC target."),
        ("5", "PCell vs SCell DL traffic share",
         "SCell bits / (PCell bits + SCell bits) = L.CA.Traffic.bits.DL.SCell / (L.CA.Traffic.bits.DL.PCell + L.CA.Traffic.bits.DL.SCell) × 100%",
         "%", "Operator identity from CA counters. On a low-band PCell + high-band SCell edge UE this share should be high."),
        ("6", "Cell-edge DL rate gain (trial KPI)",
         "(R_edge_after − R_edge_before) / R_edge_before × 100%",
         "%", "Huawei/Turkcell commercial trial reported > 30% cell-edge downlink data rate. Use the same cell-edge definition (e.g. RSRP < −110 dBm or UL SINR < 0 dB) before and after. Not a Huawei counter formula."),
        ("7", "CA UE E-RAB drop rate",
         "L.E-RAB.AbnormRel.CAUser / (L.E-RAB.AbnormRel.CAUser + L.E-RAB.NormRel.CAUser) × 100%",
         "%", "CA FPD. PCell ping-pong from SC must not raise drop."),
        ("8", "Spectrum Coordination take-effect ratio (qualitative)",
         "UEs with UL-quality PCell swap / CA UEs on the pair   —  exact counter not in repo",
         "%", "DSS FPD states this ratio falls when LTE UL RBs shrink. Track via PCell band distribution of CA UEs (high-band PCell % in center vs edge)."),
    ]
    for i, rec in enumerate(kpis):
        vals = list(rec) + [""] * 5
        r = table_row(ws, r, vals, fills=[alt_fill(i)] * 10, height=48)
        merge(ws, r - 1, 5, r - 1, COLS)

    r = subsection(ws, r, COLS, "H2.  Sample arithmetic  (cell-edge gain)")
    r = formula_box(
        ws, r, COLS,
        "Turkcell-class cell-edge DL gain  (public trial, not an MML output)",
        "Gain% = (R_after − R_before) / R_before × 100%",
        "Suppose a cell-edge CA UE (high-band PCell, UL limited) had R_before = 8 Mbit/s DL because UL ACK/NACK and UL traffic collapsed.\n"
        "After SC the same location uses low-band PCell + high-band SCell: R_after = 10.5 Mbit/s.\n"
        "Gain% = (10.5 − 8) / 8 × 100% = 31.25%.",
        "This matches the order of the published “over 30% at cell edges” result. If your after-rate does not move, first check SCell add/act success and remaining LTE UL RBs (DSS), not the 30% figure itself."
    )

    # ------------------------------------------------------------------ I
    r = section(ws, r, COLS, "I.  Licenses — FDD and TDD")
    r = headers(ws, r, ["SN", "RAT", "Feature ID", "Feature name", "Model", "Sales unit", "When consumed"] + [""] * 3)
    lic = [
        ("1", "FDD", "Not printed in CA FPD", "LTE Spectrum Coordination (LTE FDD)",
         "Not in repo — open License Control Item Lists (FDD) (CA FPD ref 91)", "See list (typically per cell)",
         "Related FPD CellCaAlgoSwitch listing names this feature on the same MO as CA bits. Do not assume a CA license (LAOFD-*) covers it."),
        ("2", "TDD", "Not printed in CA FPD", "LTE Spectrum Coordination (LTE TDD)",
         "Not in repo — open License Control Item Lists (TDD) (CA FPD ref 93)", "See list (typically per cell)",
         "Same as FDD: separate from TDLAOFD CA licenses. Enhancement (WbbCaMultiCarrierCoordSw) may share or add a WBB license — confirm in WBB FPD (CA ref 54)."),
        ("3", "FDD/TDD", "LAOFD-001001 / TDLAOFD-001001 (and 40 MHz / 3CC / 4CC+5CC as deployed)",
         "Downlink CA (this workbook Steps 1–4)", "See Step1–4 license tables", "per cell",
         "Still required. Spectrum Coordination rides on CA. License insufficiency of CA items can block CELL ACTIVATION."),
        ("4", "FDD", "MRFD160222 family (DSS) — only if DSS is also on",
         "LTE FDD and NR Flash Dynamic Spectrum Sharing", "LT1S0LFNSS00 (DSS FPD, not SC)", "per cell",
         "Listed so planners do not confuse DSS licenses with Spectrum Coordination. DSS reduces SC take-effect ratio."),
    ]
    for i, rec in enumerate(lic):
        vals = list(rec) + [""] * 3
        r = table_row(ws, r, vals, fills=[PALE_RED if i < 2 else alt_fill(i)] * 10, height=52)
        merge(ws, r - 1, 7, r - 1, COLS)
    r = note_bar(ws, r, COLS, "No Feature ID / Model for LTE Spectrum Coordination is printed in the CA FPD, DSS FPD extract, or NSA FPD extract used here. Do not copy guessed IDs into a license request. Use License Control Item Lists (FDD/TDD) and the dedicated FPD.")

    # ------------------------------------------------------------------ J
    r = section(ws, r, COLS, "J.  Deployment checklist  (after Step1 CA is green)")
    r = headers(ws, r, ["SN", "Step", "Check", "Pass criteria"] + [""] * 6)
    chk = [
        ("1", "Cluster", "High + low (or TDD high + FDD low) co-coverage pair chosen", "Overlap plot; CA neighbors exist; FREQ_MEAS_FLAG selected"),
        ("2", "CA baseline", "Step1 2CC counters non-zero for a week", "SCell add/act success high; CA UE rate > non-CA"),
        ("3", "License", "Spectrum Coordination + CA licenses on both cells", "LST LICENSE / License Control Item Lists — do not guess models"),
        ("4", "Conflicts", "DSS UL RB share, UlMultiClusterSwitch, NSA bit20, WBB load-transfer", "Documented coexistence: fewer UL RBs → fewer SC UEs; bit20 if NSA PCC must stay"),
        ("5", "HO path", "CaA5HoEventSwitch + CaA5HoEventEnhSwitch", "PCell can HO to SCell"),
        ("6", "Enable", "SpectrumCoordinationSwitch-1 on both cells of the pair", "DSP CAMGTCFG shows selected"),
        ("7", "Monitor 24–72 h", "PCell HO, SCell add/act, cell-edge DL rate, drop", "HO success stable; cell-edge DL up; no ping-pong"),
        ("8", "Optimize or roll back", "If ping-pong or drop, switch-0 and open dedicated FPD for hysteresis", "Rollback MML in section F seq 9"),
    ]
    for i, rec in enumerate(chk):
        vals = list(rec) + [""] * 6
        r = table_row(ws, r, vals, fills=[alt_fill(i)] * 10, height=32)
        merge(ws, r - 1, 4, r - 1, COLS)

    # ------------------------------------------------------------------ K
    r = section(ws, r, COLS, "K.  Sources  (every claim on this sheet traces to one of these)")
    r = headers(ws, r, ["SN", "Source", "What was taken from it"] + [""] * 7)
    src = [
        ("1", "Huawei eRAN Carrier Aggregation FPD, eRAN21.1 Issue 11 (2026-06-30), Ch.23 item 55",
         "Identifies LTE Spectrum Coordination as a related eRAN FPD. CloudAIR impact rows: WbbCaMultiCarrierCoordSw vs DlCaLbAlgoSwitch / CaLoadBalancePreAllocSwitch on WBB CA UEs/RRNs (Ch.5 and Ch.13)."),
        ("2", "CaMgtCfg.CellCaAlgoSwitch option text in related eRAN FPDs (CoMP / parameter listing)",
         "SpectrumCoordinationSwitch: enabled only if selected; PCell changes triggered based on uplink quality; FDD and TDD; default Off. Feature names “LTE Spectrum Coordination (LTE FDD/TDD)” appear on that MO."),
        ("3", "SingleRAN LTE FDD and NR Spectrum Sharing FPD (Flash DSS) function-impact table",
         "SpectrumCoordinationSwitch identity. Fewer LTE UL RBs → smaller proportion of UEs for which Spectrum Coordination takes effect. UlMultiClusterSwitch listed on the same impact."),
        ("4", "EPC-based NSA Performance Enhancement FPD",
         "RsvdSwPara6_bit20 disables Spectrum Coordination for NSA DC UEs. WbbCaMultiCarrierCoordSw = coordination enhancement. NSA DC PCC anchoring takes precedence over the enhancement."),
        ("5", "Huawei press: Turkcell Antalya, 22 Jan 2018, eRAN 13.1 commercial network",
         "DL always on high+low; UL selected high (center) or low (edge); cell-edge DL rate +30%+. CloudAIR key technology; lifts UL/DL channel binding."),
        ("6", "Huawei CloudAIR white paper (Jan 2018)",
         "Channel cloudification example: LTE spectrum coordination uses high and low bands; FDD high+low and TDD high + FDD low."),
        ("7", "This repository — not present",
         "Dedicated “LTE Spectrum Coordination” Feature Parameter Description (thresholds, MML examples, counters, license models). Obtain from 3900 & 5900 Series Base Station Product Documentation (CA FPD ref 100)."),
    ]
    for i, rec in enumerate(src):
        vals = list(rec) + [""] * 7
        r = table_row(ws, r, vals, fills=[PALE_RED if i == 6 else alt_fill(i)] * 10, height=52)
        merge(ws, r - 1, 3, r - 1, COLS)

    r = body(
        ws, r, COLS,
        "If the dedicated LTE Spectrum Coordination FPD is added to this repository later, this sheet should be extended with its trigger/leave charts, "
        "full parameter list, verbatim MML, counters, KPI, and FDD/TDD licenses — the same depth as Step1 2CC.",
        fill_hex=PALE_GOLD,
    )
    return ws
