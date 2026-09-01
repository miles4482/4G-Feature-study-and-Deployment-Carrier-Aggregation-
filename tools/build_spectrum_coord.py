#!/usr/bin/env python3
"""Sheet 7: LTE Spectrum Coordination — from the dedicated Huawei FPD.

Source: eRAN LTE Spectrum Coordination Feature Parameter Description
        eRAN21.1 Issue 01 (2025-03-10)
        Features LCOFD-131312 (FDD) / TDLCOFD-131312 (TDD)
"""
import os, sys
sys.path.insert(0, os.path.dirname(__file__))
from ca_excel_style import *
from build_step1_2cc import (
    add_param_header, add_param, add_mml_header, add_mml, add_ctr_header, add_ctr,
    FIG, COLS, _fig,
)

SRC = ("Huawei eRAN LTE Spectrum Coordination Feature Parameter Description, "
       "eRAN21.1 Issue 01 (2025-03-10)")


def build_spectrum_coord(wb):
    ws = wb.create_sheet("7. LTE Spectrum Coordination")
    setup_sheet(ws, "Spectrum Coord")
    ws.oddFooter.left.text = "Source: Huawei eRAN LTE Spectrum Coordination FPD, eRAN21.1 Issue 01 (2025-03-10)"
    set_widths(ws, [8, 28, 34, 24, 30, 56, 44, 12, 12, 12])
    r = 1
    r = banner(ws, r, COLS, "  7.  LTE Spectrum Coordination   —  FPD Ch.4 + Ch.5 Enhancement")
    r = note_bar(
        ws, r, COLS,
        SRC + ".  Applies to FDD/TDD (no FDD↔TDD principle difference, §2.4).  "
        "MML, counter IDs, license models, and Figure 4-1 are taken from this FPD.  "
        "Doc example uses downlink 3CC (LocalCellId=0/1/2).  "
        "Where the FPD does not print a factory default, Default = “See Parameter Reference”; Recommended = Table 4-1/4-2 setting or the MML sample."
    )

    # ------------------------------------------------------------------ A
    r = section(ws, r, COLS, "A.  Principal  (Ch.3 Overview + Ch.4.1 Principles)")
    r = body(
        ws, r, COLS,
        "In multi-band LTE networks, high frequencies usually encounter uplink usage limitation while downlink is still available. "
        "This limitation may even affect the deployment of high frequencies. LTE spectrum coordination lifts the downlink usage of "
        "carrier aggregation (CA) UEs at the uplink coverage edge of high frequencies.  (Ch.3)\n\n"
        "When a UE in the CA state is located at the uplink coverage edge of a high-frequency serving cell and the uplink SINR used "
        "in scheduling of the UE becomes unsatisfactory, the eNodeB evaluates whether any low-frequency SCell of the UE can provide "
        "better uplink performance. If an SCell can provide better performance, the eNodeB hands over the UE from the PCell to this "
        "SCell and simultaneously configures the original PCell as an SCell for the UE.  (Ch.4.1)\n\n"
        "The uplink SINR used in scheduling refers to that for initial selection of MCSs (see Uplink Scheduling). "
        "Figure 4-1 example: before coordination the CA UE has PCell 2600 MHz and SCell 850 MHz; after coordination UL is on 850 MHz "
        "while DL stays on 2600 MHz + 850 MHz."
    )
    r = insert_figure(ws, r, _fig("sc_fig4_1.png"), COLS,
                      "Figure 4-1  Example of LTE spectrum coordination  (source: Huawei LTE Spectrum Coordination FPD p.6)  —  2600 MHz PCell + 850 MHz SCell")

    r = subsection(ws, r, COLS, "A1.  Enablement conditions  (all must be met)  —  Ch.4.1")
    r = headers(ws, r, ["SN", "Condition", "MO / parameter", "Required setting", "Notes"] + [""] * 5)
    en = [
        ("1", "MLB admission of unnecessary HO must be off",
         "CellMlbHo.MlbMatchOtherFeatureMode → HoAdmitSwitch", "Deselected",
         "On the original PCells and SCells of the UEs (Table 4-1)."),
        ("2", "SCell configuration during handover",
         "ENodeBAlgoSwitch.CaAlgoSwitch → HoWithSccCfgSwitch", "Selected",
         "Prerequisite function (Table 4-1 / §4.3.2.1)."),
        ("3", "Spectrum coordination master switch",
         "CaMgtCfg.CellCaAlgoSwitch → SpectrumCoordinationSwitch", "Selected",
         "On the original PCells and SCells of the UEs (Table 4-1)."),
    ]
    for i, rec in enumerate(en):
        vals = list(rec) + [""] * 5
        r = table_row(ws, r, vals, fills=[alt_fill(i)] * 10, height=32)
        merge(ws, r - 1, 5, r - 1, COLS)

    r = subsection(ws, r, COLS, "A2.  Options that MUST be selected after the function is enabled  (Ch.4.1)")
    r = bullets(ws, r, COLS, [
        "ENodeBAlgoSwitch.CaAlgoExtSwitch → CaEnhancedPreAllocSwitch  AND  CellAlgoSwitch.PucchAlgoSwitch → Dl2CCAckResShareSw  —  “Selecting these options prevents the traffic volume in low-frequency cells from increasing.”",
        "ENodeBAlgoSwitch.CaAlgoSwitch → SccA2RmvSwitch  —  “Selecting this option prevents the possibility that UEs do not report the measurement results of SCells or the strongest intra-frequency neighboring cells of SCells.”",
    ])
    ws.row_dimensions[r - 1].height = 56

    r = subsection(ws, r, COLS, "A3.  CA band combination / UL 2CC compatibility  (Ch.4.1)")
    r = bullets(ws, r, COLS, [
        "Compatible with all CA band combinations for downlink 2CC to 8CC aggregation.",
        "Does NOT take effect for UEs that have two CCs aggregated for BOTH uplink and downlink, because uplink SCells will not be selected as target serving cells.",
        "For CA UEs with 2 CCs in UL and n CCs in DL: if n = 2 → Spectrum Coordination does not take effect; if 2 < n ≤ 8 → it takes effect.",
        "UEs must support CA (§4.3.4).",
    ])
    ws.row_dimensions[r - 1].height = 72

    r = subsection(ws, r, COLS, "A4.  Recommended live-network scenario  (Ch.4.2.1 Benefits)")
    r = body(
        ws, r, COLS,
        "Suitable for CA-enabled multi-band networks. A larger spacing between high and low frequencies brings greater benefits.\n"
        "Scenario 1 — both high- and low-frequency carriers already deployed: percentage of PUSCH MCSs 0 to 3 > 10% in high-frequency cells; "
        "uplink interference to low-frequency cells minus that to high-frequency cells < 5 dB; CA UE penetration > 25%. "
        "Network User Downlink Average Throughput increases. If the CA UE proportion stays unchanged, overall CA UE average DL throughput increases; "
        "if that proportion rises, the increase shrinks and CA UE average DL throughput may even drop.\n"
        "Scenario 2 — only high-frequency deployed, enable this when a low-frequency carrier is being deployed: both User DL and User UL Average Throughput increase."
    )

    # ------------------------------------------------------------------ B
    r = section(ws, r, COLS, "B.  Triggering conditions  (Ch.4.1 Inter-Frequency Handover)")
    r = note_bar(
        ws, r, COLS,
        "Trigger is uplink SINR used in scheduling (initial MCS selection), compared with CaMgtCfg.SpectrumCoordSinrThld.  "
        "Doc optimization MML sets SpectrumCoordSinrThld = −2 on every original PCell.  "
        "This is NOT CA event A1–A6 as the trigger (A4/A2 are used afterwards to pick the target SCell and to manage SCells)."
    )

    r = formula_box(
        ws, r, COLS,
        "B1.  UL-SINR entering condition  (Ch.4.1 + Table 4-2 + optimization MML)",
        "When a CA UE is transmitting uplink data in its uplink serving cell AND\n"
        "    (uplink SINR used in scheduling)  <  CaMgtCfg.SpectrumCoordSinrThld\n"
        "the inter-frequency handover and carrier-management procedures of Ch.4.1 start.\n"
        "SINR is the one used for initial MCS selection (Uplink Scheduling FPD).",
        "Doc MML: MOD CAMGTCFG: LocalCellId=0/1/2, SpectrumCoordSinrThld=-2;\n"
        "Sample UE at high-frequency PCell: UL SINR used in scheduling = −3 dB, threshold = −2 dB.\n"
        "Compare: −3 < −2  → TRUE.",
        "LTE Spectrum Coordination evaluation STARTS. If UL SINR = 0 dB: 0 < −2 is FALSE → no SC handover from this trigger."
    )

    r = subsection(ws, r, COLS, "B2.  Target-cell (new PCell) evaluation  —  a downlink SCell is suitable only if BOTH are true")
    r = bullets(ws, r, COLS, [
        "Downlink signal strength of that SCell reaches the RSRP threshold for event A4  AND  is higher than all of its intra-frequency neighboring cells. If either fails, the SCell is not a target.",
        "Recommendation (Ch.4.1): CaGroupCell.PCellA4RsrpThd (group) or PccFreqCfg.PccA4RsrpThd (adaptive)  >  InterFreqHoGroup.InterFreqHoA4ThdRsrp, to prevent a decrease in handover success rate.",
        "The SCell must also meet PCell conditions (see Carrier Aggregation FPD — this workbook Step1).",
        "If no downlink SCell meets the conditions: eNodeB does not initiate an inter-frequency handover.",
        "If exactly one candidate: handover to that cell.",
        "If two or more candidates: selection rules in B3.",
        "After HO: original SCell becomes PCell; original PCell becomes SCell. In adaptive CA (not group-based), carrier management (section C) then runs.",
    ])
    ws.row_dimensions[r - 1].height = 130

    r = formula_box(
        ws, r, COLS,
        "B2 worked example  —  is this SCell a legal SC target?  (A4 + strongest intra-freq)",
        "SCell is suitable iff  (SCell RSRP reaches CA event A4)  AND  (SCell RSRP > every intra-frequency neighbor of that SCell).\n"
        "CA A4 entering (this workbook / CA FPD Ch.4.3):  Mn + Ofn + Ocn − Hys  >  Thresh\n"
        "Thresh = CarrAggrA4ThdRsrp + SCellA4Offset/SccA4Offset  (offsets 0 in the CA FPD MML examples).",
        "Assume adaptive CA, PccA4RsrpThd = −105 dBm (CA FPD MML), offsets = 0, Hys = 1 dB.\n"
        "Low-frequency SCell Mn = −98 dBm, Ofn=Ocn=0.  Mn+Ofn+Ocn−Hys = −99 dB.  −99 > −105 → A4 TRUE.\n"
        "Strongest intra-freq neighbor of that SCell = −108 dBm.  SCell −98 > −108 → strongest TRUE.",
        "SCell IS a candidate. If Mn = −112 dBm: −113 is NOT > −105 → SCell is not a target; if no other candidate, no SC handover."
    )

    r = subsection(ws, r, COLS, "B3.  Choosing among multiple candidate SCells  (Ch.4.1)")
    r = headers(ws, r, ["Order", "Intelligent selection OFF (CaSmartSelectionSwitch)", "Intelligent selection ON"] + [""] * 7)
    sel = [
        ("1", "Highest PCell priority (CaGroupCell.PreferredPCellPriority) or PCC priority (PccFreqCfg.PreferredPccPriority)",
         "Highest downlink air interface capability (see CA FPD)"),
        ("2", "If still tied: largest bandwidth",
         "If still tied: highest PCell / PCC priority (same MOs as left)"),
        ("3", "If still tied: lowest center frequency",
         "If still tied: larger bandwidth; then lowest center frequency"),
    ]
    for i, rec in enumerate(sel):
        vals = list(rec) + [""] * 7
        r = table_row(ws, r, vals, fills=[alt_fill(i)] * 10, height=40)
        merge(ws, r - 1, 3, r - 1, COLS)

    r = subsection(ws, r, COLS, "B4.  Worked multi-candidate example  (doc selection rules + 3CC MML EARFCNs)")
    r = body(
        ws, r, COLS,
        "Doc 3CC optimization MML uses three downlink EARFCNs 123 / 567 / 890 as PCC/SCC pairs. Sample (intelligent selection OFF):\n"
        "UE PCell = high-frequency cell on EARFCN 890, UL SINR = −3 dB < −2 dB → SC starts. Two DL SCells both pass A4+strongest:\n"
        "  Cell A  EARFCN 123, PreferredPccPriority=2, BW=10 MHz\n"
        "  Cell B  EARFCN 567, PreferredPccPriority=2, BW=20 MHz\n"
        "Same PCC priority → pick largest bandwidth → Cell B. After HO: PCell = Cell B, original 890-cell becomes SCell. "
        "If both had BW=20 MHz, pick the lower EARFCN (123).",
        fill_hex=EXAMPLE,
    )
    ws.row_dimensions[r - 1].height = 88

    # ------------------------------------------------------------------ C
    r = section(ws, r, COLS, "C.  Leaving / carrier management after the PCell change  (Ch.4.1 Carrier Management)")
    r = note_bar(ws, r, COLS, "Runs when the UE is camping on the new PCell AND CA is adaptive (not group-based). Controlled by SccFreqCfg.SccA2RsrpThldExtendedOfs. Doc MML sets this to −10 on every PCC↔SCC pair of the 3CC example.")

    r = subsection(ws, r, COLS, "C1.  SccA2RsrpThldExtendedOfs = 0")
    r = body(ws, r, COLS,
             "eNodeB delivers A2 measurement configurations related to SCells. After an A2 report that contains an SCell, "
             "eNodeB sends RRC Connection Reconfiguration to remove that SCell. Same CA A2 procedure as Carrier Aggregation FPD.")

    r = formula_box(
        ws, r, COLS,
        "C2.  SccA2RsrpThldExtendedOfs ≠ 0  and intelligent selection ON  (Ch.4.1 steps i–iv)",
        "First A2 threshold  =  CarrAggrA2ThdRsrp + SccA2Offset\n"
        "On first A2: evaluate a NEW A2 entering condition with\n"
        "    New A2 threshold  =  CarrAggrA2ThdRsrp + SccA2Offset + SccA2RsrpThldExtendedOfs\n"
        "If RSRP meets the new condition → remove SCell (end).\n"
        "If not → delete previous A2 cfg; deliver new A2 cfg + A1 cfg for this SCell.\n"
        "    A1 threshold  =  CarrAggrA4ThdRsrp + SccA4Offset\n"
        "Before further A2/A1, intelligent selection may pick a better SCell.\n"
        "Further A2 containing this SCell → remove it.  A1 containing this SCell → release new A2, restore previous A2 (threshold back to CarrAggrA2ThdRsrp + SccA2Offset).",
        "Doc MML: SccA2RsrpThldExtendedOfs = −10. Assume CarrAggrA2ThdRsrp = −115 dBm, SccA2Offset = 0, CarrAggrA4ThdRsrp = −105, SccA4Offset = 0 (CA-style numbers; A2/A4 bases are CA FPD).\n"
        "First A2 thresh = −115 dBm. New A2 thresh = −115 + 0 + (−10) = −125 dBm.\n"
        "SCell RSRP in the report = −118 dBm.\n"
        "−118 meets first A2 (below −115) but does NOT meet new A2 (need to be below −125).",
        "SCell is NOT removed on the first A2. eNodeB switches to the extended A2 + A1 pair (A1 thresh = −105 dBm) and intelligent selection may replace the SCell. "
        "If a later A2 still contains this SCell, it is removed."
    )

    r = subsection(ws, r, COLS, "C3.  SccA2RsrpThldExtendedOfs ≠ 0  and intelligent selection OFF")
    r = formula_box(
        ws, r, COLS,
        "Single A2 threshold (no two-step A1)",
        "A2 threshold  =  CarrAggrA2ThdRsrp + SccA2Offset + SccA2RsrpThldExtendedOfs\n"
        "After an A2 report containing an SCell, eNodeB removes that SCell.",
        "Same numbers: threshold = −115 + 0 + (−10) = −125 dBm. SCell RSRP = −118 dBm.\n"
        "−118 is NOT below −125.",
        "No SCell removal on this report. If RSRP later falls to −128 dBm, A2 fires and the SCell is removed."
    )

    r = subsection(ws, r, COLS, "C4.  Operator rollback  (Ch.4.4.1.2 Deactivation MML)")
    r = body(ws, r, COLS,
             "MOD CAMGTCFG: LocalCellId=0/1/2, CellCaAlgoSwitch=SpectrumCoordinationSwitch-0;  "
             "stops UL-SINR-triggered PCell swaps. CA itself is unchanged. Doc example is the 3CC set of three cells.")

    r = subsection(ws, r, COLS, "C5.  Network impacts of the extra HOs  (Ch.4.2.2)")
    r = bullets(ws, r, COLS, [
        "KPI change after HO depends on whether the original SCell had better or worse channel quality than the original PCell. Better → UL/DL cell throughput, IBLER, average MCS, RRC re-establishment rate, drop rate improve; worse → they deteriorate. CQI changes.",
        "PCell/SCell CA user and bit counters change: L.Traffic.User.PCell.DL.Avg, L.Traffic.User.SCell.DL.Avg, L.CA.Traffic.bits.DL.PCell, L.CA.Traffic.bits.DL.SCell.",
        "More handovers → fewer UEs in DRX. For voice UEs, MOS probably decreases.",
        "With HoWithSccCfgSwitch selected, SCell add success rate is affected by HO success rate: L.CA.DLSCell.Add.Succ / L.CA.DLSCell.Add.Att.",
    ])
    ws.row_dimensions[r - 1].height = 88

    # ------------------------------------------------------------------ D
    r = section(ws, r, COLS, "D.  Prerequisite / exclusive / impacted features  (Ch.4.3.2 + Ch.5.3.2)  —  one feature per row")
    r = headers(ws, r, ["SN", "Relation", "RAT", "Feature", "Related parameter(s)", "Impact / rule", "Notes"] + [""] * 3)
    feats = [
        ("1", "Prerequisite", "FDD/TDD", "Carrier aggregation",
         "CA switches vary by scenario — see this workbook Steps 1–4",
         "Must already be active. Compatible with DL 2CC–8CC. Not for UL2CC+DL2CC UEs.",
         "§4.3.2.1 / §5.3.2.1"),
        ("2", "Prerequisite", "FDD/TDD", "SCell configuration during handovers",
         "ENodeBAlgoSwitch.CaAlgoSwitch → HoWithSccCfgSwitch",
         "Must be selected. After enablement, SCell add success is tied to HO success.",
         "§4.3.2.1; also required for WBB enhancement."),
        ("3", "Exclusive", "FDD/TDD", "Admission control of MLB triggering over unnecessary handovers",
         "CellMlbHo.MlbMatchOtherFeatureMode → HoAdmitSwitch",
         "Must be deselected on original PCells and SCells or SC does not enable.",
         "§4.3.2.2  Mobility Management in Connected Mode / Intra-RAT MLB"),
        ("4", "Must-select with SC", "FDD/TDD", "Enhanced pre-allocation + PUCCH ACK resource share",
         "ENodeBAlgoSwitch.CaAlgoExtSwitch → CaEnhancedPreAllocSwitch; CellAlgoSwitch.PucchAlgoSwitch → Dl2CCAckResShareSw",
         "Must be selected after SC is enabled to prevent low-frequency cell traffic from increasing.",
         "Ch.4.1 + Table 4-1"),
        ("5", "Must-select with SC", "FDD/TDD", "SCell A2 remove",
         "ENodeBAlgoSwitch.CaAlgoSwitch → SccA2RmvSwitch",
         "Must be selected so UEs report SCell / strongest intra-freq neighbor measurements.",
         "Ch.4.1 + Table 4-1"),
        ("6", "Impact", "FDD", "Uplink full-antenna reception (massive MIMO)",
         "SectorSplitGroup.SectorSplitSwitch → UL_COVERAGE_BOOST_SW",
         "After it takes effect, proportion of CA UEs at UL coverage edge of massive MIMO carriers decreases → possible change in SC inter-freq HO count.",
         "§4.3.2.3  Massive MIMO Enhancements (FDD)"),
        ("7", "Impact (carrier mgmt off)", "FDD/TDD", "SCC coverage threshold adaptation",
         "MultiCarrUnifiedSch.MultiCarrierUnifiedSchSw → SCC_ADAPT_COV_THLD_SW",
         "When enabled, carrier management in LTE spectrum coordination does not take effect.",
         "§4.3.2.3  Multi-carrier Unified Scheduling"),
        ("8", "Impact", "FDD", "LTE FDD and NR Flash DSS",
         "SpectrumCloud.SpectrumCloudSwitch = LTE_NR_SPECTRUM_SHR",
         "Fewer LTE UL RBs → affects the proportion of UEs for which SC takes effect.",
         "§4.3.2.3"),
        ("9", "Impact", "FDD", "LTE FDD and NR Uplink Spectrum Sharing (SUL DSS)",
         "SpectrumCloud.SpectrumCloudSwitch = LTE_NR_UPLINK_SPECTRUM_SHR",
         "Fewer LTE UL RBs → same take-effect proportion impact.",
         "§4.3.2.3"),
        ("10", "Impact", "FDD", "DSS and NR Flexible Refarming",
         "SpectrumCloud.SpectrumCloudEnhSwitch → DSS_FLEX_REFARM_SW",
         "After LTE sharing carriers are shut down, fewer cells can be selected for SC → may take effect for fewer UEs.",
         "§4.3.2.3"),
        ("11", "Impact / prevent HO", "FDD/TDD", "NSA PCC anchoring",
         "NsaDcMgmtConfig.NsaDcAlgoSwitch → NSA_PCC_ANCHORING_SWITCH; NsaDcMgmtConfig.NsaDcUeLteFunActivationSw → SPCT_COORD_INTER_FREQ_HO_SW",
         "SC may HO an NSA UE from an NSA PCC anchor to a non-anchor carrier. To prevent: deselect SPCT_COORD_INTER_FREQ_HO_SW so SC does not take effect for NSA UEs.",
         "§4.3.2.3  EPC-based NSA Performance Enhancement"),
        ("12", "Enhancement prerequisite", "FDD/TDD", "Specified service carrier (WBB)",
         "Cell.SpecifiedCellFlag = MBBSERCELL",
         "Required for WBB/RRN enhancement: low-frequency cells that do not allow WBB access must be MBB service cells.",
         "§5.3.2.1"),
        ("13", "Enhancement prerequisite", "FDD/TDD", "CA for out-of-band relay",
         "See Relay FPD (scenario-dependent switches)",
         "Required only for RRN enhancement.",
         "§5.3.2.1"),
        ("14", "Enhancement exclusive", "FDD/TDD", "Carrier-level rate control for CA UEs",
         "CellAlgoSwitch.DacqEnhancementSwitch → CaUserLimitOptSwitch",
         "Mutually exclusive with LTE spectrum coordination enhancement.",
         "§5.3.2.2  Rate Control Based on User Types"),
        ("15", "Enhancement vs NSA", "FDD/TDD", "NSA PCC anchoring vs WbbCaMultiCarrierCoordSw",
         "NsaDcMgmtConfig.NsaDcAlgoSwitch → NSA_PCC_ANCHORING_SWITCH; CaMgtCfg.CellCaAlgoSwitch → WbbCaMultiCarrierCoordSw",
         "NSA PCC anchoring takes precedence. If both enabled, only NSA PCC anchoring takes effect.",
         "§5.3.2.3"),
        ("16", "Enhancement impact", "TDD", "Inter-CC load transfer (cell load / load differences)",
         "ENodeBAlgoSwitch.CaLbAlgoSwitch → DlCaLbAlgoSwitch; ENodeBAlgoSwitch.CaAlgoSwitch → CaLoadBalancePreAllocSwitch",
         "Does not take effect for WBB CA UEs or RRNs on which enhancement has taken effect.",
         "§5.3.2.3  (matches CA FPD CloudAIR rows)"),
    ]
    for rec in feats:
        kind = rec[1]
        fh = PALE_GREEN if "Prerequisite" in kind or "Must-select" in kind else PALE_RED if "Exclusive" in kind else PALE_ORANGE if "Impact" in kind else alt_fill(int(rec[0]))
        r = table_row(ws, r, list(rec) + [""] * 3, fills=[fh] * 10,
                      bolds=[False, True, False, True, False, False, False], height=58)
        merge(ws, r - 1, 7, r - 1, COLS)

    # ------------------------------------------------------------------ E
    r = section(ws, r, COLS, "E.  Parameters  (Tables 4-1, 4-2, 5-9, 5-10 + related MOs in Ch.4.1/5.1)")
    r = add_param_header(ws, r)
    r = add_param(ws, r, 1, "CaMgtCfg", "CellCaAlgoSwitch / SpectrumCoordinationSwitch",
                  "See Parameter Reference", "Select on original PCells and SCells (Table 4-1)",
                  "Master enable for LTE Spectrum Coordination. Function is enabled only when this bit, HoWithSccCfgSwitch, and HoAdmitSwitch-deselected are all met.",
                  "Doc MML: CellCaAlgoSwitch=SpectrumCoordinationSwitch-1 on LocalCellId=0,1,2 (3CC example).")
    r = add_param(ws, r, 2, "CaMgtCfg", "SpectrumCoordSinrThld",
                  "See Parameter Reference", "−2 (doc optimization MML, original PCells)",
                  "Uplink SINR threshold used in scheduling (initial MCS selection). When UL SINR of the uplink serving cell falls below this, SC inter-frequency HO + carrier management start.",
                  "Table 4-2: set to recommended value for original PCells. Related: Uplink Scheduling FPD.", related=True)
    r = add_param(ws, r, 3, "ENodeBAlgoSwitch", "CaAlgoSwitch / HoWithSccCfgSwitch",
                  "See Parameter Reference", "Select (Table 4-1 and 5-9)",
                  "SCell configuration during handovers. Prerequisite for SC and for WBB enhancement.",
                  "Related: SCell add success vs HO success (Ch.4.2.2).", related=True)
    r = add_param(ws, r, 4, "ENodeBAlgoSwitch", "CaAlgoSwitch / SccA2RmvSwitch",
                  "See Parameter Reference", "Select (Table 4-1, after SC is enabled)",
                  "Prevents UEs from not reporting SCell / strongest intra-freq neighbor measurements.",
                  "Must be selected after the function is enabled (Ch.4.1).", related=True)
    r = add_param(ws, r, 5, "ENodeBAlgoSwitch", "CaAlgoExtSwitch / CaEnhancedPreAllocSwitch",
                  "See Parameter Reference", "Select (Table 4-1, after SC is enabled)",
                  "Together with Dl2CCAckResShareSw, prevents traffic volume in low-frequency cells from increasing.",
                  "Related: CellAlgoSwitch.PucchAlgoSwitch / Dl2CCAckResShareSw.", related=True)
    r = add_param(ws, r, 6, "CellAlgoSwitch", "PucchAlgoSwitch / Dl2CCAckResShareSw",
                  "See Parameter Reference", "Select on original PCells and SCells (Table 4-1)",
                  "PUCCH ACK resource sharing. Paired with CaEnhancedPreAllocSwitch as above.",
                  "Doc MML on LocalCellId=0,1,2.", related=True)
    r = add_param(ws, r, 7, "CellMlbHo", "MlbMatchOtherFeatureMode / HoAdmitSwitch",
                  "See Parameter Reference", "Deselect on original PCells and SCells (Table 4-1)",
                  "Admission control of MLB triggering over unnecessary handovers. Mutually exclusive with SC — must be off.",
                  "Doc MML: MlbMatchOtherFeatureMode=HoAdmitSwitch-0.", related=True)
    r = add_param(ws, r, 8, "SccFreqCfg", "SccA2RsrpThldExtendedOfs",
                  "See Parameter Reference", "−10 (doc optimization MML, every PCC↔SCC pair)",
                  "Extended offset relative to the RSRP threshold for event A2 that controls SCell removal after the SC PCell change (adaptive CA). 0 = simple A2 remove; non-zero = extended procedure (C2/C3).",
                  "Table 4-2 / 5-10. Related: CarrAggrA2ThdRsrp, SccA2Offset, CarrAggrA4ThdRsrp, SccA4Offset, CaSmartSelectionSwitch.", related=True)
    r = add_param(ws, r, 9, "CaGroupCell / PccFreqCfg", "PCellA4RsrpThd / PccA4RsrpThd",
                  "See Parameter Reference (CA FPD MML −105)", "Greater than InterFreqHoA4ThdRsrp (Ch.4.1 recommendation)",
                  "A4 RSRP used to decide whether a DL SCell is a legal SC target. Group vs adaptive MO.",
                  "Related: PCellA4RsrqThd / PccA4RsrqThd for WBB A5 Thresh2.", related=True)
    r = add_param(ws, r, 10, "InterFreqHoGroup", "InterFreqHoA4ThdRsrp",
                  "See Parameter Reference", "Keep below PCellA4/PccA4 RSRP (Ch.4.1)",
                  "Coverage HO A4. If SC A4 is not higher, HO success rate may decrease.",
                  "Related to SC target evaluation, not the UL-SINR trigger.", related=True)
    r = add_param(ws, r, 11, "CaGroupCell / PccFreqCfg", "PreferredPCellPriority / PreferredPccPriority",
                  "See Parameter Reference", "Plan so the preferred low-frequency coverage cell wins when UL SINR is poor",
                  "First tie-break among multiple SC target SCells when intelligent selection is OFF (and second tie-break when ON).",
                  "Then largest bandwidth, then lowest center frequency.", related=True)
    r = add_param(ws, r, 12, "ENodeBAlgoSwitch", "CaAlgoSwitch / CaSmartSelectionSwitch",
                  "See Parameter Reference", "Site — changes target-cell and A2-extended procedures",
                  "Intelligent selection of serving cell combinations. Changes B3 selection order and C2 vs C3 A2 handling.",
                  "See this workbook sheet 2.", related=True)
    r = add_param(ws, r, 13, "CaMgtCfg", "CellCaAlgoSwitch / WbbCaMultiCarrierCoordSw",
                  "See Parameter Reference", "Select on low- and high-frequency cells for enhancement (Table 5-9)",
                  "LTE spectrum coordination enhancement (WBB UEs and RRNs). WBB: also needs HoWithSccCfgSwitch and MBBSERCELL. RRN: MBBSERCELL + this bit.",
                  "Ch.5. Exclusive with CaUserLimitOptSwitch. NSA PCC anchoring takes precedence.", related=True)
    r = add_param(ws, r, 14, "Cell", "SpecifiedCellFlag",
                  "See Parameter Reference", "MBBSERCELL on low-frequency cells (Table 5-9)",
                  "Marks low-frequency cells that do not allow WBB access as MBB service cells. Required for WBB/RRN enhancement.",
                  "After enhancement, WBB CA UEs can HO to MBB service cells and then aggregate WBB+MBB.", related=True)
    r = add_param(ws, r, 15, "CellWttxParaCfg", "WbbMultiCarrierCoordA2Ofs",
                  "See Parameter Reference", "2 (doc optimization MML, low-frequency cells)",
                  "WBB multi-carrier coordination A2 offset. Enters the WBB A2 threshold and A5 Thresh2 formulas (Tables 5-1, 5-2, 5-3).",
                  "Table 5-10. Related: InterFreqHoA2ThdRsrp/Rsrq, SpidCfg.InterFreqHoA2RsrpThdFactor / RsrqThdFactor.", related=True)
    r = add_param(ws, r, 16, "NsaDcMgmtConfig", "NsaDcUeLteFunActivationSw / SPCT_COORD_INTER_FREQ_HO_SW",
                  "See Parameter Reference", "Deselect to stop SC for NSA UEs (Ch.4.3.2.3)",
                  "If SC would HO an NSA UE from an NSA PCC anchor to a non-anchor, deselect this option so SC does not take effect for NSA UEs.",
                  "Related: NSA_PCC_ANCHORING_SWITCH (takes precedence over enhancement).", related=True)

    # ------------------------------------------------------------------ F
    r = section(ws, r, COLS, "F.  MML  (verbatim from Ch.4.4.1.2 and Ch.5.4.1.2)")
    r = note_bar(ws, r, COLS, "Ch.4 example is downlink 3CC (LocalCellId=0,1,2; EARFCN 123/567/890). Ch.5 example: low-frequency LocalCellId=0 as MBBSERCELL, EARFCN 123 PCC + 456 SCC. Replace IDs with live-network values.")
    r = add_mml_header(ws, r)
    mmls = [
        (1, "1", "Both", "FDD/TDD", "Activate SC",
         "MOD CELLMLBHO: LocalCellId=0, MlbMatchOtherFeatureMode=HoAdmitSwitch-0;",
         "Doc: Enabling LTE spectrum coordination (taking downlink 3CC as an example). Repeat 0,1,2."),
        (2, "2", "Both", "FDD/TDD", "Activate SC",
         "MOD CELLMLBHO: LocalCellId=1, MlbMatchOtherFeatureMode=HoAdmitSwitch-0;",
         "Original PCell/SCell set."),
        (3, "3", "Both", "FDD/TDD", "Activate SC",
         "MOD CELLMLBHO: LocalCellId=2, MlbMatchOtherFeatureMode=HoAdmitSwitch-0;",
         "Original PCell/SCell set."),
        (4, "4", "Both", "FDD/TDD", "Activate SC",
         "MOD CELLALGOSWITCH: LocalCellId=0, PucchAlgoSwitch=Dl2CCAckResShareSw-1;",
         "Prevents low-frequency traffic increase (with CaEnhancedPreAllocSwitch)."),
        (5, "5", "Both", "FDD/TDD", "Activate SC",
         "MOD CELLALGOSWITCH: LocalCellId=1, PucchAlgoSwitch=Dl2CCAckResShareSw-1;",
         ""),
        (6, "6", "Both", "FDD/TDD", "Activate SC",
         "MOD CELLALGOSWITCH: LocalCellId=2, PucchAlgoSwitch=Dl2CCAckResShareSw-1;",
         ""),
        (7, "7", "Both", "FDD/TDD", "Activate SC",
         "MOD ENODEBALGOSWITCH: CaAlgoSwitch=HoWithSccCfgSwitch-1&SccA2RmvSwitch-1, CaAlgoExtSwitch=CaEnhancedPreAllocSwitch-1;",
         "Doc single eNodeB-level command (HO-with-SCell + A2 remove + enhanced pre-alloc)."),
        (8, "8", "Both", "FDD/TDD", "Activate SC",
         "MOD CAMGTCFG: LocalCellId=0, CellCaAlgoSwitch=SpectrumCoordinationSwitch-1;",
         "Master bit on each of the three cells."),
        (9, "9", "Both", "FDD/TDD", "Activate SC",
         "MOD CAMGTCFG: LocalCellId=1, CellCaAlgoSwitch=SpectrumCoordinationSwitch-1;",
         ""),
        (10, "10", "Both", "FDD/TDD", "Activate SC",
         "MOD CAMGTCFG: LocalCellId=2, CellCaAlgoSwitch=SpectrumCoordinationSwitch-1;",
         ""),
        (11, "11", "Both", "FDD/TDD", "Optimize SC",
         "MOD CAMGTCFG: LocalCellId=0, SpectrumCoordSinrThld=-2;",
         "Doc: Setting the SINR threshold (3CC example). Repeat 0,1,2."),
        (12, "12", "Both", "FDD/TDD", "Optimize SC",
         "MOD CAMGTCFG: LocalCellId=1, SpectrumCoordSinrThld=-2;",
         ""),
        (13, "13", "Both", "FDD/TDD", "Optimize SC",
         "MOD CAMGTCFG: LocalCellId=2, SpectrumCoordSinrThld=-2;",
         ""),
        (14, "14", "Adaptive", "FDD/TDD", "Optimize SC",
         "MOD SCCFREQCFG: PccDlEarfcn=123, SccDlEarfcn=567, SccA2RsrpThldExtendedOfs=-10;",
         "Doc: extended A2 offset on every PCC↔SCC pair of {123,567,890}."),
        (15, "15", "Adaptive", "FDD/TDD", "Optimize SC",
         "MOD SCCFREQCFG: PccDlEarfcn=123, SccDlEarfcn=890, SccA2RsrpThldExtendedOfs=-10;",
         ""),
        (16, "16", "Adaptive", "FDD/TDD", "Optimize SC",
         "MOD SCCFREQCFG: PccDlEarfcn=567, SccDlEarfcn=123, SccA2RsrpThldExtendedOfs=-10;",
         ""),
        (17, "17", "Adaptive", "FDD/TDD", "Optimize SC",
         "MOD SCCFREQCFG: PccDlEarfcn=567, SccDlEarfcn=890, SccA2RsrpThldExtendedOfs=-10;",
         ""),
        (18, "18", "Adaptive", "FDD/TDD", "Optimize SC",
         "MOD SCCFREQCFG: PccDlEarfcn=890, SccDlEarfcn=123, SccA2RsrpThldExtendedOfs=-10;",
         ""),
        (19, "19", "Adaptive", "FDD/TDD", "Optimize SC",
         "MOD SCCFREQCFG: PccDlEarfcn=890, SccDlEarfcn=567, SccA2RsrpThldExtendedOfs=-10;",
         ""),
        (20, "20", "Both", "FDD/TDD", "Deactivate SC",
         "MOD CAMGTCFG: LocalCellId=0, CellCaAlgoSwitch=SpectrumCoordinationSwitch-0;",
         "Doc deactivation (3CC). Repeat 1 and 2."),
        (21, "21", "Both", "FDD/TDD", "Deactivate SC",
         "MOD CAMGTCFG: LocalCellId=1, CellCaAlgoSwitch=SpectrumCoordinationSwitch-0;",
         ""),
        (22, "22", "Both", "FDD/TDD", "Deactivate SC",
         "MOD CAMGTCFG: LocalCellId=2, CellCaAlgoSwitch=SpectrumCoordinationSwitch-0;",
         ""),
        (23, "23", "Both", "FDD/TDD", "Activate enhancement",
         "MOD ENODEBALGOSWITCH: CaAlgoSwitch=HoWithSccCfgSwitch-1;",
         "Ch.5.4.1.2 Activating LTE spectrum coordination enhancement."),
        (24, "24", "Both", "FDD/TDD", "Activate enhancement",
         "MOD CELL: LocalCellId=0, SpecifiedCellFlag=MBBSERCELL;",
         "Low-frequency cell as MBB service cell."),
        (25, "25", "Both", "FDD/TDD", "Activate enhancement",
         "MOD CAMGTCFG: LocalCellId=0, CellCaAlgoSwitch=WbbCaMultiCarrierCoordSw-1;",
         "Low- and high-frequency cells (doc 0 and 1)."),
        (26, "26", "Both", "FDD/TDD", "Activate enhancement",
         "MOD CAMGTCFG: LocalCellId=1, CellCaAlgoSwitch=WbbCaMultiCarrierCoordSw-1;",
         ""),
        (27, "27", "Both", "FDD/TDD", "Optimize enhancement",
         "MOD CELLWTTXPARACFG: LocalCellId=0, WbbMultiCarrierCoordA2Ofs=2;",
         "Doc: Setting WbbMultiCarrierCoordA2Ofs (repeat 0,1)."),
        (28, "28", "Both", "FDD/TDD", "Optimize enhancement",
         "MOD CELLWTTXPARACFG: LocalCellId=1, WbbMultiCarrierCoordA2Ofs=2;",
         ""),
        (29, "29", "Adaptive", "FDD/TDD", "Optimize enhancement",
         "MOD SCCFREQCFG: PccDlEarfcn=123, SccDlEarfcn=456, SccA2RsrpThldExtendedOfs=-10;",
         "Doc: low-frequency PCC EARFCN 123, high-frequency WBB SCC EARFCN 456."),
        (30, "30", "Both", "FDD/TDD", "Deactivate enhancement",
         "MOD CAMGTCFG: LocalCellId=0, CellCaAlgoSwitch=WbbCaMultiCarrierCoordSw-0;",
         "Doc deactivation. Repeat LocalCellId=1."),
        (31, "31", "Both", "FDD/TDD", "Deactivate enhancement",
         "MOD CAMGTCFG: LocalCellId=1, CellCaAlgoSwitch=WbbCaMultiCarrierCoordSw-0;",
         ""),
    ]
    for rec in mmls:
        r = add_mml(ws, r, *rec)

    # ------------------------------------------------------------------ G
    r = section(ws, r, COLS, "G.  Counters  (Tables 4-3, 4-4, 5-4–5-8, 5.4.2)")
    r = add_ctr_header(ws, r)
    ctrs = [
        (1, "1526729994", "L.HHO.InterFreq.ULquality.PrepAttOut", "UL-quality inter-freq HO preparations", "Table 4-3 activation: values increase ⇒ SC has taken effect"),
        (2, "1526729996", "L.HHO.InterFreq.ULquality.ExecAttOut", "UL-quality inter-freq HO executions", "Table 4-3"),
        (3, "1526729998", "L.HHO.InterFreq.ULquality.ExecSuccOut", "UL-quality inter-freq HO success", "Table 4-3; success rate KPI"),
        (4, "1526729995", "L.HHO.InterFddTdd.ULquality.PrepAttOut", "UL-quality FDD↔TDD HO preparations", "Table 4-3"),
        (5, "1526729997", "L.HHO.InterFddTdd.ULquality.ExecAttOut", "UL-quality FDD↔TDD HO executions", "Table 4-3"),
        (6, "1526729999", "L.HHO.InterFddTdd.ULquality.ExecSuccOut", "UL-quality FDD↔TDD HO success", "Table 4-3"),
        (7, "1526728426", "L.Traffic.User.PCell.DL.Avg", "Avg CA UEs using this cell as PCell", "Table 4-4 / 5-8; Ch.4.2.2 says this moves after PCell swap"),
        (8, "1526728427", "L.Traffic.User.SCell.DL.Avg", "Avg CA UEs using this cell as SCell", "Table 4-4 / 5-8"),
        (9, "1526732658", "L.CA.Traffic.bits.DL.PCell", "DL bits on PCell of CA UEs", "Table 4-4 / 5-8"),
        (10, "1526729259", "L.CA.Traffic.bits.DL.SCell", "DL bits on SCell of CA UEs", "Table 4-4 / 5-8"),
        (11, "1526728564", "L.Thrp.bits.DL.CAUser", "CA UE DL bits", "Table 4-4 / 5-5 / 5-8; CA UE rate KPI"),
        (12, "1526740440", "L.Thrp.bits.DL.LastTTI.CAUser", "CA UE last-TTI DL bits", "Table 4-4 / 5-5"),
        (13, "1526740441", "L.Thrp.Time.DL.RmvLastTTI.CAUser", "CA UE DL time last-TTI removed", "Table 4-4 / 5-5"),
        (14, "1526729045", "L.CA.DLSCell.Add.Att", "DL SCell add attempts", "Table 4-4 / 5-8; SCell cfg success vs HO"),
        (15, "1526729046", "L.CA.DLSCell.Add.Succ", "DL SCell add success", "Table 4-4 / 5-8"),
        (16, "1526746999–7008", "L.Thrp.DL.BitRate.Samp.CaUe.Index0 … Index9", "CA UE DL bit-rate samples (10 bins)", "Table 4-4 monitoring"),
        (17, "1526728261", "L.Thrp.bits.DL", "All-UE DL bits", "Table 5-4 MBB perceived rate"),
        (18, "1526729005", "L.Thrp.bits.DL.LastTTI", "All-UE last-TTI DL bits", "Table 5-4"),
        (19, "1526745879", "L.Thrp.bits.DL.WBB", "WBB DL bits", "Tables 5-4 / 5-6"),
        (20, "1526745880", "L.Thrp.bits.DL.LastTTI.WBB", "WBB last-TTI DL bits", "Tables 5-4 / 5-6"),
        (21, "1526729015", "L.Thrp.Time.DL.RmvLastTTI", "All-UE DL time last-TTI removed", "Table 5-4"),
        (22, "1526745881", "L.Thrp.Time.DL.RmvLastTTI.WBB", "WBB DL time last-TTI removed", "Tables 5-4 / 5-6; Ch.5.2.2 may increase"),
        (23, "1526735542", "L.Thrp.Relay.bits.DL", "RRN DL bits", "Table 5-7"),
        (24, "1526735543", "L.Thrp.Relay.Time.DL", "RRN DL time", "Table 5-7; Ch.5.4.2: increase on low-freq cells ⇒ RRN enhancement took effect"),
        (25, "1526745954", "L.ChMeas.PRB.DL.PDSCH.WBBUsed.Avg", "Avg PDSCH PRBs used by WBB", "Ch.5.2.2: decreases on low-freq, increases on high-freq"),
        (26, "1526740478", "L.Traffic.SpecSerUser.Avg", "Specified-service (MBB) users", "Ch.5.4.2: decrease to a non-zero value on low-freq cells ⇒ WBB enhancement took effect"),
    ]
    for rec in ctrs:
        r = add_ctr(ws, r, *rec)

    # ------------------------------------------------------------------ H
    r = section(ws, r, COLS, "H.  KPI formulas  (Ch.4.2.1, 4.2.2, 4.4.2, 5.2)")
    r = headers(ws, r, ["SN", "KPI", "Formula (from document)", "Unit", "Notes"] + [""] * 5)
    kpis = [
        ("1", "UL-quality inter-freq HO success",
         "L.HHO.InterFreq.ULquality.ExecSuccOut / L.HHO.InterFreq.ULquality.ExecAttOut × 100%",
         "%", "From Table 4-3 counters. Rise in Prep/Exec/Succ verifies activation."),
        ("2", "UL-quality FDD↔TDD HO success",
         "L.HHO.InterFddTdd.ULquality.ExecSuccOut / L.HHO.InterFddTdd.ULquality.ExecAttOut × 100%",
         "%", "Table 4-3"),
        ("3", "Average downlink throughput of CA UEs",
         "(L.Thrp.bits.DL.CAUser − L.Thrp.bits.DL.LastTTI.CAUser) / L.Thrp.Time.DL.RmvLastTTI.CAUser",
         "bit/s", "Ch.4.2.1 Scenario 1. If CA UE proportion rises after SC, this KPI may even drop."),
        ("4", "SCell configuration success rate",
         "L.CA.DLSCell.Add.Succ / L.CA.DLSCell.Add.Att × 100%",
         "%", "Ch.4.2.2 / 5.2.2: affected by HO success once HoWithSccCfgSwitch is on."),
        ("5", "Perceived data rate of MBB UEs (enhancement)",
         "[(L.Thrp.bits.DL − L.Thrp.bits.DL.LastTTI) − (L.Thrp.bits.DL.WBB − L.Thrp.bits.DL.LastTTI.WBB)] / (L.Thrp.Time.DL.RmvLastTTI − L.Thrp.Time.DL.RmvLastTTI.WBB)",
         "bit/s", "Ch.5.2.1 Table 5-4. Expected to increase when enhancement is activated (low-freq were normal cells)."),
        ("6", "Perceived data rate of CA UEs (enhancement impact)",
         "(L.Thrp.bits.DL.CAUser − L.Thrp.bits.DL.LastTTI.CAUser) / L.Thrp.Time.DL.RmvLastTTI.CAUser",
         "bit/s", "Ch.5.2.2: overall average perceived CA UE rate decreases after enhancement."),
        ("7", "Perceived data rate of WBB UEs",
         "(L.Thrp.bits.DL.WBB − L.Thrp.bits.DL.LastTTI.WBB) / L.Thrp.Time.DL.RmvLastTTI.WBB",
         "bit/s", "Ch.5.2.2 Table 5-6: decreases. CA takes effect for a smaller proportion of WBB UEs."),
        ("8", "Perceived data rate of RRNs",
         "L.Thrp.Relay.bits.DL / L.Thrp.Relay.Time.DL",
         "bit/s", "Ch.5.2.2 Table 5-7: decreases."),
    ]
    for i, rec in enumerate(kpis):
        vals = list(rec) + [""] * 5
        r = table_row(ws, r, vals, fills=[alt_fill(i)] * 10, height=48)
        merge(ws, r - 1, 5, r - 1, COLS)

    r = formula_box(
        ws, r, COLS,
        "H worked example  —  CA UE average DL throughput  (Ch.4.2.1 formula + sample counts)",
        "R = (L.Thrp.bits.DL.CAUser − L.Thrp.bits.DL.LastTTI.CAUser) / L.Thrp.Time.DL.RmvLastTTI.CAUser",
        "Suppose a 15-minute bin on the high/low pair after SC:\n"
        "  L.Thrp.bits.DL.CAUser = 7.20×10^12 bit\n"
        "  L.Thrp.bits.DL.LastTTI.CAUser = 0.45×10^12 bit\n"
        "  L.Thrp.Time.DL.RmvLastTTI.CAUser = 1.50×10^8 ms\n"
        "  R = (7.20 − 0.45)×10^12 / 1.50×10^8 = 6.75×10^12 / 1.50×10^8 = 4.50×10^4 bit/ms = 45.0 Mbit/s.",
        "Compare the same formula on cell-edge bins (PUSCH MCS 0–3 share, Ch.4.2.1 Scenario 1) before vs after SC. "
        "If CA UE proportion also rose, Ch.4.2.1 warns this average can fall even when network User DL Average Throughput rose."
    )

    # ------------------------------------------------------------------ I
    r = section(ws, r, COLS, "I.  Licenses — FDD and TDD  (Ch.4.3.1 / 5.3.1  —  same license covers enhancement)")
    r = headers(ws, r, ["SN", "RAT", "Feature ID", "Feature name", "Model", "Sales unit", "When consumed"] + [""] * 3)
    lic = [
        ("1", "FDD", "LCOFD-131312", "LTE Spectrum Coordination (LTE FDD)", "LT1SUNPLTE00", "per cell",
         "Ch.4.3.1 and Ch.5.3.1. Enhancement uses the same license. Insufficiency of certain items affects CELL ACTIVATION — see License Control Item Lists."),
        ("2", "TDD", "TDLCOFD-131312", "LTE Spectrum Coordination (LTE TDD)", "LT1SLTESPC00", "per cell",
         "Ch.4.3.1 and Ch.5.3.1. Same note on cell activation. No FDD/TDD feature-difference in switches or principles (§2.4)."),
        ("3", "FDD/TDD", "(this workbook Steps 1–4)", "Carrier aggregation (prerequisite)", "LAOFD / TDLAOFD / LEOFD as deployed", "per cell",
         "CA must already be licensed and active. SC does not replace CA licenses."),
    ]
    for i, rec in enumerate(lic):
        vals = list(rec) + [""] * 3
        r = table_row(ws, r, vals, fills=[alt_fill(i)] * 10, height=48)
        merge(ws, r - 1, 7, r - 1, COLS)

    # ------------------------------------------------------------------ J  Enhancement
    r = section(ws, r, COLS, "J.  LTE Spectrum Coordination Enhancement  (Ch.5)  —  WBB UEs and RRNs")
    r = subsection(ws, r, COLS, "J1.  WBB UEs  (Ch.5.1.1)  —  trigger is downlink RSRP, not UL SINR")
    r = body(
        ws, r, COLS,
        "In addition to Ch.4, enhancement applies to CA-capable WBB UEs. When the UE is at the coverage edge of its high-frequency serving cell "
        "and has low downlink RSRP there, the eNodeB evaluates whether a low-frequency cell can provide better performance; if yes, HO PCell to that "
        "low-frequency cell and configure the original PCell as SCell.\n"
        "WBB UEs are identified by SPID or by QCI (see WBB FPD).\n"
        "Enabled when ALL of: (1) low-frequency cells that do not allow WBB access are MBBSERCELL; (2) HoWithSccCfgSwitch selected; "
        "(3) WbbCaMultiCarrierCoordSw selected. Takes effect only for WBB CA UEs that entered non-MBB service cells by initial access.\n"
        "HO is triggered only if the UE is not an emergency call, not eMBMS, and has no QCI 1/65/66 bearer. "
        "If SC-enhancement A2 and coverage-based A2 are both met, coverage-based A2 takes precedence and enhancement does not take effect.\n"
        "After HO to an MBB service cell: original PCell becomes SCell; WBB-dedicated cells can be SCells; MBB service cells cannot be SCells; "
        "eNodeB will not remove SCells based on A2; DL data is scheduled only in SCells whenever possible (if all SCell config fails, UE may use the MBB PCell)."
    )
    ws.row_dimensions[r - 1].height = 130

    r = formula_box(
        ws, r, COLS,
        "J2.  WBB A2 threshold  (Table 5-1)  +  A5 Thresh1/Thresh2  (Tables 5-2 / 5-3)",
        "A2 (SPID WBB) = InterFreqHoA2ThdRsrp + SpidCfg.InterFreqHoA2RsrpThdFactor + WbbMultiCarrierCoordA2Ofs\n"
        "A2 (QCI WBB)  = InterFreqHoA2ThdRsrp + WbbMultiCarrierCoordA2Ofs\n"
        "Then eNodeB delivers A5; IntraRatHoComm.InterFreqHoA4TrigQuan = RSRP / RSRQ / BOTH.\n"
        "A5 Thresh1 is fixed: RSRP −43 dBm, RSRQ −3 dB (same constants as CA FPD).\n"
        "A5 Thresh2 (example, adaptive + RSRP + QCI WBB) =\n"
        "    Max{ PccFreqCfg.PccA4RsrpThd ,  InterFreqHoA2ThdRsrp + WbbMultiCarrierCoordA2Ofs + 2 }",
        "Doc MML: WbbMultiCarrierCoordA2Ofs = 2. Assume InterFreqHoA2ThdRsrp = −105 dBm, PccA4RsrpThd = −105 dBm, QCI-identified WBB UE, adaptive, RSRP.\n"
        "A2 thresh = −105 + 2 = −103 dBm.\n"
        "A5 Thresh1 = −43 dBm (fixed).  A5 Thresh2 = Max{−105, −105+2+2} = Max{−105, −101} = −101 dBm.",
        "PCell RSRP must stay below A2 (−103) throughout TTT; neighbor must exceed A5 Thresh2 (−101) with Thresh1 always true at realistic RSRP. "
        "If the A5 target is an MBB service cell that cannot aggregate with any current serving cell, eNodeB does not HO."
    )

    r = subsection(ws, r, COLS, "J3.  RRNs  (Ch.5.1.2)")
    r = bullets(ws, r, COLS, [
        "Out-of-band relay CA: if RRN PCell is not an MBB service cell, MBB cells cannot be SCells. If PCell is MBB, non-MBB cells can be SCells; DL scheduled only in SCells when possible; if all SCell config fails, RRN may use the MBB cell.",
        "Enabled when MBBSERCELL is set on the low-frequency no-WBB-access cells AND WbbCaMultiCarrierCoordSw is selected.",
        "Verification (Ch.5.4.2): L.Thrp.Relay.Time.DL (1526735543) increases on low-frequency cells.",
    ])
    ws.row_dimensions[r - 1].height = 64

    # ------------------------------------------------------------------ K
    r = section(ws, r, COLS, "K.  Hardware, UE, MAE, verification  (Ch.4.3.3–4.4.3, 5.3.3–5.4.3)")
    r = bullets(ws, r, COLS, [
        "SC (Ch.4): FDD 3900/5900 macro; TDD 3900/5900 macro + DBS3900/DBS5900 LampSite. LBBPc boards cannot be used. RF: no requirements.",
        "Enhancement (Ch.5): 3900/5900 macro + LampSite. TDD: cells on LBBPc cannot be used. FDD boards: no requirements. RF: no requirements.",
        "UEs must support CA.",
        "MAE-Deployment: Feature Operation and Maintenance (fast batch) or Feature Configuration Using the MAE-Deployment. OSS GUI may differ from the interactive video.",
        "Activation proof (Ch.4.4.2): Table 4-3 UL-quality HO counters increase. Enhancement (Ch.5.4.2): L.Traffic.SpecSerUser.Avg on low-freq cells decreases to non-zero (WBB); L.Thrp.Relay.Time.DL on low-freq cells increases (RRN).",
        "Huawei note (Ch.2.1): this FPD is activation guidance. Gains depend on the live scenario; contact Huawei professional service for optimal gains.",
    ])
    ws.row_dimensions[r - 1].height = 110
    return ws
