# Step2 3CC, Step3 4CC, Step4 5CC — same section order as Step1.
import os, sys
sys.path.insert(0, os.path.dirname(__file__))
from ca_excel_style import *
from build_step1_2cc import (
    add_param_header, add_param, add_mml_header, add_mml, add_ctr_header, add_ctr,
    peak_chart, trigger_leave_block, FIG, COLS, _fig,
)


def _kpi_header(ws, r):
    return headers(ws, r, ["SN", "KPI", "Formula (from document)", "Unit", "Notes"] + [""] * 5)


def _lic_header(ws, r):
    return headers(ws, r, ["SN", "RAT", "Feature ID", "Feature name", "Model", "Sales unit", "When consumed"] + [""] * 3)


def build_3cc(wb):
    ws = wb.create_sheet("4. Step2 3CC Config")
    setup_sheet(ws, "3CC")
    set_widths(ws, [8, 28, 32, 22, 28, 55, 42, 12, 12, 12])
    r = 1
    r = banner(ws, r, COLS, "  Step 2 —  Downlink 3CC configuration   (document Ch.6)   —  same structure as Step1")
    r = note_bar(ws, r, COLS, "Prerequisite: Downlink 2CC (Step1) must already be active. Then turn on CaDl3CCSwitch on each possible PCell. "
                 "Group: add ≥3 cells and candidate SCells per PCell. Adaptive: ≥2 candidate SCCs per PCC. "
                 "Triggering / leaving procedures are the same Ch.4.6 state machine as 2CC (eNodeB prefers the maximum supported CC count).")

    r = section(ws, r, COLS, "A.  Principal  (Ch.6.1–6.2)")
    r = insert_figure(ws, r, _fig("emb_p227_0.png"), COLS, "Figure 6-1  Downlink 3CC aggregation  (source: Huawei FPD p.214)")
    r = body(ws, r, COLS,
             "Aggregates three intra- or inter-band carriers. Works intra-eNB, inter-eNB coordination, and relaxed backhaul.\n"
             "Table 6-1 switches:\n"
             "• Group-based, BW ≤ 60 MHz: CaDl3CCSwitch ON. CaDl3CCExtSwitch not required.\n"
             "• Adaptive FDD, BW ≤ 40 MHz: CaDl3CCSwitch ON.\n"
             "• Adaptive FDD, 40 < BW ≤ 60 MHz: CaDl3CCSwitch ON AND CaDl3CCExtSwitch ON (this extra bit is FDD-only; TDD 3CC is not affected — Ch.2.4).\n"
             "• Adaptive TDD, BW ≤ 60 MHz: CaDl3CCSwitch ON only.\n"
             "If CqiAdaptiveCfg.SimulAckNackAndCqiFmt3Sw = ON, HARQ ACK/NACK and periodic CQI for the three CCs can be multiplexed on PUCCH format 3; if OFF they cannot share a subframe.\n"
             "PUCCH format-3 overhead: adaptive PUCCH (PucchSwitch ON) spares +1 RB (PUSCH RBs drop by ≥1, must be multiple of 2/3/5) → UL throughput down. "
             "Fixed PUCCH: FDD converts 1 CQI RB to format 3; TDD converts 2 CQI RBs → more aperiodic CQI, slight DL rate drop. PCell DL IBLER fluctuates.")
    ws.row_dimensions[r - 1].height = 110
    r = peak_chart(ws, r, "Theoretical peak data rates for DL 3CC  (Table 6-2, Mbit/s)",
                   ["2x2 + 64QAM", "2x2 + 256QAM", "4x4 + 64QAM", "4x4 + 256QAM"],
                   [449.3, 587.4, 899.6, 1175.0],
                   [336, 429, 651, 858], "3cc")

    r = trigger_leave_block(ws, r, "3CC")

    r = section(ws, r, COLS, "D.  Prerequisite, mutually exclusive, and mutually impacted features")
    r = headers(ws, r, ["SN", "Relation", "RAT", "Feature", "Parameter(s)", "Required action", "Impact / note"] + [""] * 3)
    rows = [
        ("1", "Prerequisite", "FDD/TDD", "Downlink 2CC aggregation", "None extra (Step1 must be ON)", "Complete Step1 first", "Ch.6.3.3.1"),
        ("2", "Exclusive", "FDD", "PUCCH measurement", "CellAlgoSwitch.PucchAlgoSwitch → PucchMeasOptSwitch  must be OFF", "Deactivate PUCCH measurement (SFN doc)", "Ch.6.3.3.2"),
        ("3", "Impact", "FDD", "UMTS and LTE Zero Bufferzone", "ULZeroBufferZone.ZeroBufZoneSwitch → UMTS_LTE_ZERO_BUFFER_ZONE_SW", "Do not use 5/10 MHz bufferzone cell as PCell (few PUSCH/SRS)", "Ch.6.3.3.3"),
        ("4", "Impact", "FDD", "UMTS and LTE Spectrum Sharing (LTE FDD)", "SpectrumCloud.SpectrumCloudSwitch = UL_SPECTRUM_SHARING", "5 MHz cells not recommended as PCell (PUCCH so large SRS cannot be configured)", ""),
        ("5", "Impact", "FDD", "UMTS/LTE SS based on DC-HSDPA", "SpectrumCloud.SpectrumCloudSwitch = DC_HSDPA_BASED_UL_SPECTRUM_SHR", "Same 5 MHz PCell warning", "Related to row 4 (same MO, different enum)", True),
        ("6", "Impact", "FDD", "LTE FDD and NR Flash DSS", "SpectrumCloud.SpectrumCloudSwitch = LTE_NR_SPECTRUM_SHR", "If LTE preferential resource % = 0, LTE UEs cannot access", ""),
        ("7", "Impact", "FDD", "Hybrid DSS based on asymmetric BW", "SpectrumCloudSwitch = LTE_NR_SPECTRUM_SHR AND SpectrumCloud.SpectrumCloudEnhSwitch → LNR_SPECTRUM_SHR_ASYM_SW", "Same: LTE resource % = 0 → no LTE access", "Related to row 6", True),
        ("8", "Impact", "TDD", "SRS resource migration", "UlInterfSuppressCfg.RemoteIntrfDlEnhSwitch → REMOTE_INTRF_BF_ENH_SW", "When both conditions met, SRS migration WINS: 3CC stops; when migration stops, 3CC resumes", ""),
        ("9", "Constraint", "FDD/TDD", "UE / EPC", "UE Rel-12+; MBR ≥ Table 6-2 peak", "Rel-10 is not enough for 3CC", "Ch.6.3.6"),
        ("10", "Inherit", "FDD/TDD", "All Step1 exclusive features still apply", "Super combined cell, >100 km, in-band relay, MBSFN shutdown, MIE energy saving", "Keep them OFF", "2CC exclusives remain"),
    ]
    for rec in rows:
        related = rec[-1] is True
        rec = rec[:-1] if rec[-1] is True else rec
        vals = list(rec) + [""] * 3
        rel = rec[1]
        fh = PALE_GREEN if rel == "Prerequisite" else (PALE_RED if rel == "Exclusive" else (PALE_GOLD if related or rel in ("Constraint", "Inherit") else alt_fill(int(rec[0]))))
        r = table_row(ws, r, vals, fills=[fh] * 10, height=40)
        merge(ws, r - 1, 7, r - 1, COLS)

    r = section(ws, r, COLS, "E.  All parameters  (3CC-specific on top of Step1 set)")
    r = note_bar(ws, r, COLS, "Prepare basic data exactly as Step1 (Ch.5.4.1.1). Then set Table 6-5 / 6-6 / 6-7. Related 2CC parameters that 3CC still needs are listed immediately under the 3CC switch.")
    r = add_param_header(ws, r)
    r = add_param(ws, r, 1, "CaMgtCfg", "CellCaAlgoSwitch / CaDl3CCSwitch", "OFF", "1 on every possible PCell",
                  "Master switch for downlink 3CC.",
                  "Parent 3CC switch. Group or adaptive. Related children below.")
    r = add_param(ws, r, 2, "CaMgtCfg", "CellCaAlgoSwitch / CaDl3CCExtSwitch", "OFF", "1 on each cell of a 40–60 MHz FDD adaptive set",
                  "FDD adaptive only: 3CC in 60 MHz window.",
                  "Related to CaDl3CCSwitch (same MO). Consumes LAOFD-080208. TDD does not use this bit.", related=True)
    r = add_param(ws, r, 3, "CaMgtCfg", "CellCaAlgoSwitch / CaDl2CCExtSwitch", "See Step1", "Keep as in Step1 if 2CC 40 MHz window still needed",
                  "2CC 40 MHz bit remains required for the 2CC layer.",
                  "Related prerequisite Step1 switch.", related=True)
    r = add_param(ws, r, 4, "CqiAdaptiveCfg", "SimulAckNackAndCqiFmt3Sw", "See Parameter Reference", "ON to multiplex ACK/NACK + periodic CQI on format 3",
                  "3CC HARQ + CQI multiplexing.", "If OFF, they cannot be sent in the same subframe.")
    r = add_param(ws, r, 5, "CaGroup / PccFreqCfg / SccFreqCfg", "(all Step1 objects)", "As Step1", "Group: ≥3 cells. Adaptive: ≥2 SCCs per PCC",
                  "Topology must support 3 CCs.", "Related: SCellPriority / SccPriority order decides which two SCells are chosen.", related=True)
    r = add_param(ws, r, 6, "CaMgtCfg", "Forbid3ccUeRatioThld", "100", "80 (doc 2CC optional MML, still valid)",
                  "UE-proportion to forbid 3CC or higher.", "Must be < Forbid2ccUeRatioThld. Related Forbid2/4/5.", related=True)
    r = add_param(ws, r, 7, "RlcPdcpParaGroup", "(see 4CC/5CC timers)", "N/A for 3CC FPD", "Not required by Ch.6",
                  "Dl4cc5cc timers are introduced at 4CC, not 3CC.", "Listed so operators do not apply them too early.")

    r = section(ws, r, COLS, "F.  MML commands  (Ch.6.4.1.2 / 6.4.2.2  —  after Step1 MML)")
    r = add_mml_header(ws, r)
    cmds = [
        (1, "0", "Both", "FDD/TDD", "Prerequisite",
         "(Run ALL Step1 MML first: neighbors, group or PCC/SCC, EnhancedPccAnchorSwitch, CA_UE_MULTI_CC_STAT_OPT_SW, ACT CELL.)",
         "Ch.6: “Before activating this function, add at least three cells / configure at least two candidate SCCs”."),
        (2, "1.1", "Group", "FDD", "Activation",
         'MOD CAMGTCFG: LocalCellId=0, CellCaAlgoSwitch=CaDl3CCSwitch-1;',
         "Ch.6.4.1.2.1. Repeat for every possible PCell. Deactivate: CaDl3CCSwitch-0."),
        (3, "1.2", "Adaptive", "FDD", "Activation",
         'MOD CAMGTCFG: LocalCellId=0, CellCaAlgoSwitch=CaDl3CCSwitch-1;',
         "Ch.6.4.1.2.2."),
        (4, "1.3", "Adaptive", "FDD", "Optional 60 MHz",
         'MOD CAMGTCFG: LocalCellId=0, CellCaAlgoSwitch=CaDl3CCExtSwitch-1;',
         "When aggregated BW of the three CCs is 40 < BW ≤ 60 MHz. Doc also sets LocalCellId=1 and 2."),
        (5, "1.4", "Adaptive", "FDD", "Optional 60 MHz",
         'MOD CAMGTCFG: LocalCellId=1, CellCaAlgoSwitch=CaDl3CCExtSwitch-1;',
         ""),
        (6, "1.5", "Adaptive", "FDD", "Optional 60 MHz",
         'MOD CAMGTCFG: LocalCellId=2, CellCaAlgoSwitch=CaDl3CCExtSwitch-1;',
         ""),
        (7, "2.1", "Group", "TDD", "Activation",
         'MOD CAMGTCFG: LocalCellId=0, CellCaAlgoSwitch=CaDl3CCSwitch-1;',
         "Ch.6.4.2.2.1 after TDD group of ≥3 cells."),
        (8, "2.2", "Adaptive", "TDD", "Activation",
         'MOD CAMGTCFG: LocalCellId=0, CellCaAlgoSwitch=CaDl3CCSwitch-1;',
         "Ch.6.4.2.2.2. TDD has no CaDl3CCExtSwitch."),
        (9, "9", "Both", "FDD/TDD", "Deactivation",
         'MOD CAMGTCFG: LocalCellId=0, CellCaAlgoSwitch=CaDl3CCSwitch-0;',
         "Leaves 2CC running if Step1 switches stay ON."),
    ]
    for rec in cmds:
        r = add_mml(ws, r, *rec)

    r = section(ws, r, COLS, "G.  Counter list  (Ch.6.4.3 Table 6-8 + Ch.6.4.4 Table 6-9  —  plus all Step1 counters)")
    r = add_ctr_header(ws, r)
    for rec in [
        (1, "1526732907", "L.Traffic.User.PCell.DL.3CC.Avg", "Avg UEs in DL 3CC (this cell as PCell)", "Activation verify (any Table 6-8 ≠ 0)"),
        (2, "1526732908", "L.Traffic.User.PCell.DL.3CC.Max", "Max UEs in DL 3CC", "Activation + monitoring"),
        (3, "1526732915", "L.Traffic.User.PCell.DL.3CC.Active.Avg", "Avg UEs with 3 carriers active", "MULTI_CC_STAT_OPT_SW changes stepping"),
        (4, "1526732916", "L.Traffic.User.PCell.DL.3CC.Active.Max", "Max UEs with 3 carriers active", "Same stepping rule"),
        (5, "1526732917", "L.CA.DL.PCell.3CC.Act.Dur", "Total DL active duration of 3CC UEs", "Monitoring"),
        (6, "1526737809", "L.Thrp.Time.DL.3CC.CAUser", "DL PDCP duration of 3CC UEs", "3CC throughput denominator"),
        (7, "1526733012", "L.Thrp.bits.DL.3CC.CAUser", "DL PDCP bits of 3CC UEs", "3CC throughput numerator"),
        (8, "—", "(all Step1 counters)", "Still required: add/act/deact/HO/drop/PRB", "Ch.6.4.4: “In addition to the counters listed in 5.4.3”"),
    ]:
        r = add_ctr(ws, r, *rec)

    r = section(ws, r, COLS, "H.  KPI formulas for 3CC")
    r = _kpi_header(ws, r)
    for i, rec in enumerate([
        ("1", "Throughput of UEs in DL 3CC state",
         "L.Thrp.bits.DL.3CC.CAUser / L.Thrp.Time.DL.3CC.CAUser", "bit/s",
         "Ch.6.4.4. MULTI_CC_STAT_OPT_SW ON: counted as long as 3 carriers are active (even if a 4th configured SCell is inactive)."),
        ("2", "3CC active user (avg)", "L.Traffic.User.PCell.DL.3CC.Active.Avg", "UE", "Verify feature ON"),
        ("3", "Reuse all Step1 KPIs", "Drop, HO, add/act success — same formulas as Step1 sheet H", "%", "3CC UEs are a subset of CA UEs"),
    ]):
        vals = list(rec) + [""] * 5
        r = table_row(ws, r, vals, fills=[alt_fill(i)] * 10, height=36)
        merge(ws, r - 1, 5, r - 1, COLS)

    r = section(ws, r, COLS, "I.  Licenses — FDD and TDD  (Ch.6.3.1 / 6.3.2)")
    r = _lic_header(ws, r)
    for i, rec in enumerate([
        ("1", "FDD", "LAOFD-080207", "Carrier Aggregation for Downlink 3CC in 40MHz", "LT1SCAD40M00", "per cell",
         "Adaptive: each PCC-freq cell with CaDl3CCSwitch ON + each CA cell on its SCC freqs. Group: ≥3 inter-freq FDD cells AND (CaDl3CCSwitch on any cell OR CaDl4CCSwitch on any TDD cell in the group) → every group cell consumes one unit."),
        ("2", "FDD", "LAOFD-080208", "Carrier Aggregation for Downlink 3CC in 60MHz", "LT1SCAD60M00", "per cell",
         "Adaptive: cell with CaDl3CCExtSwitch ON. Group: ≥3 inter-freq FDD, 3CC/4CC switch as above, AND a 3-cell combination > 40 MHz total BW."),
        ("3", "TDD", "TDLAOFD-081405", "Carrier Aggregation for Downlink 3CC", "LT1SCAD3CC00", "per cell",
         "Every TDD cell involved in DL 3CC."),
        ("4", "Both", "(Step1 licenses)", "LTE-A Introduction + 2CC 40MHz if applicable", "See Step1", "per cell",
         "Still required. 3CC does not replace 2CC licenses."),
    ]):
        vals = list(rec) + [""] * 3
        r = table_row(ws, r, vals, fills=[alt_fill(i)] * 10, height=48)
        merge(ws, r - 1, 7, r - 1, COLS)

    r = section(ws, r, COLS, "J.  Hardware / networking / verification")
    r = bullets(ws, r, COLS, [
        "Hardware, RF, networking: same as Step1 (Ch.6.3.4–6.3.5 point to 5.3.4–5.3.5).",
        "UE Rel-12+. Verify Table 6-8 ≠ 0. Message tracing same IEs as Step1 (sCellToAddModList with two SCells).",
        "Inter-eNB: still prepare Ch.16 or Ch.17 data if used.",
    ])
    return ws


def build_4cc(wb):
    ws = wb.create_sheet("5. Step3 4CC Config")
    setup_sheet(ws, "4CC")
    set_widths(ws, [8, 28, 32, 22, 28, 55, 42, 12, 12, 12])
    r = 1
    r = banner(ws, r, COLS, "  Step 3 —  Downlink 4CC configuration   (document Ch.7)   —  same structure as Step1")
    r = note_bar(ws, r, COLS, "Prerequisites: Downlink 2CC AND Downlink 3CC (CaDl3CCSwitch). Then CaDl4CCSwitch. "
                 "Group: ≥4 cells. Adaptive: ≥3 candidate SCCs per PCC. Same Ch.4.6 triggering/leaving as Step1.")

    r = section(ws, r, COLS, "A.  Principal  (Ch.7.1–7.2)")
    r = insert_figure(ws, r, _fig("emb_p241_0.png"), COLS, "Figure 7-1  Downlink 4CC aggregation  (source: Huawei FPD p.228)")
    r = body(ws, r, COLS,
             "Aggregates four intra- or inter-band carriers. Controlled by CaDl4CCSwitch on CaMgtCfg.CellCaAlgoSwitch.\n"
             "PUCCH format-3 simultaneous ACK/NACK + CQI: CqiAdaptiveCfg.SimulAckNackAndCqiFmt3Sw ON/OFF as for 3CC.\n"
             "PUCCH overhead / IBLER impact: same pattern as 3CC ( +1 PUSCH RB loss if adaptive PUCCH; FDD 1 CQI RB / TDD 2 CQI RBs converted if fixed PUCCH).")
    r = peak_chart(ws, r, "Theoretical peak data rates for DL 4CC  (Table 7-1, Mbit/s)",
                   ["2x2 + 64QAM", "2x2 + 256QAM", "4x4 + 64QAM", "4x4 + 256QAM"],
                   [599.1, 783.3, 1199.4, 1566.6],
                   [448, 572, 868, 1144], "4cc")

    r = trigger_leave_block(ws, r, "4CC")

    r = section(ws, r, COLS, "D.  Prerequisite, mutually exclusive, and mutually impacted features")
    r = headers(ws, r, ["SN", "Relation", "RAT", "Feature", "Parameter(s)", "Required action", "Impact / note"] + [""] * 3)
    for rec in [
        ("1", "Prerequisite", "FDD/TDD", "Downlink 2CC", "None extra", "Step1 ON", "Ch.7.3.3.1"),
        ("2", "Prerequisite", "FDD/TDD", "Downlink 3CC", "CaMgtCfg.CellCaAlgoSwitch → CaDl3CCSwitch", "Step2 ON", "Related to CaDl4CCSwitch (same MO)", True),
        ("3", "Exclusive", "FDD", "PUCCH measurement", "CellAlgoSwitch.PucchAlgoSwitch → PucchMeasOptSwitch OFF", "Deactivate", "Ch.7.3.3.2"),
        ("4", "Impact", "FDD", "UMTS/LTE Zero Bufferzone", "ULZeroBufferZone.ZeroBufZoneSwitch → UMTS_LTE_ZERO_BUFFER_ZONE_SW", "Avoid 5/10 MHz bufferzone PCell", ""),
        ("5", "Impact", "FDD", "UMTS/LTE Spectrum Sharing", "SpectrumCloud.SpectrumCloudSwitch = UL_SPECTRUM_SHARING", "5 MHz not recommended as PCell", ""),
        ("6", "Impact", "FDD", "UMTS/LTE SS based on DC-HSDPA", "SpectrumCloudSwitch = DC_HSDPA_BASED_UL_SPECTRUM_SHR", "5 MHz not recommended as PCell", "Related to row 5", True),
        ("7", "Impact", "TDD", "SRS resource migration", "UlInterfSuppressCfg.RemoteIntrfDlEnhSwitch → REMOTE_INTRF_BF_ENH_SW", "SRS migration pre-empts 4CC; 4CC resumes when migration stops", ""),
        ("8", "Constraint", "FDD/TDD", "UE / EPC", "UE Rel-12+; MBR ≥ Table 7-1", "", ""),
        ("9", "Inherit", "FDD/TDD", "Step1 exclusives", "Super combined cell, >100 km, in-band relay, MBSFN, MIE ES", "Keep OFF", ""),
    ]:
        related = rec[-1] is True
        rec = rec[:-1] if rec[-1] is True else rec
        vals = list(rec) + [""] * 3
        rel = rec[1]
        fh = PALE_GREEN if rel == "Prerequisite" else (PALE_RED if rel == "Exclusive" else (PALE_GOLD if related or rel in ("Constraint", "Inherit") else alt_fill(int(rec[0]))))
        r = table_row(ws, r, vals, fills=[fh] * 10, height=36)
        merge(ws, r - 1, 7, r - 1, COLS)

    r = section(ws, r, COLS, "E.  All parameters  (4CC-specific on top of Step1/Step2)")
    r = add_param_header(ws, r)
    r = add_param(ws, r, 1, "CaMgtCfg", "CellCaAlgoSwitch / CaDl4CCSwitch", "OFF", "1 on every possible PCell",
                  "Master switch for downlink 4CC.", "Parent. Requires CaDl3CCSwitch already 1.")
    r = add_param(ws, r, 2, "CaMgtCfg", "CellCaAlgoSwitch / CaDl3CCSwitch", "1 after Step2", "Keep 1",
                  "3CC prerequisite bit on the same MO.", "Related parent/child on CellCaAlgoSwitch.", related=True)
    r = add_param(ws, r, 3, "RlcPdcpParaGroup", "Dl4cc5ccUeReorderingTimer", "See Parameter Reference", "Treordering_m15 (doc Table 7-5)",
                  "DL beyond-3CC UE reordering timer.",
                  "Valid only if RlcPdcpParaGroup.RlcMode = RlcMode_AM. Related: RlcMode must be AM.", related=True)
    r = add_param(ws, r, 4, "RlcPdcpParaGroup", "Dl4cc5ccUeStatProhTimer", "See Parameter Reference", "m15 (doc Table 7-5)",
                  "DL beyond-3CC UE status-prohibit timer.",
                  "Same AM-mode constraint. Related to reordering timer (same MO).", related=True)
    r = add_param(ws, r, 5, "RlcPdcpParaGroup", "RlcMode", "Site", "RlcMode_AM for the timers to apply",
                  "RLC mode of the group.", "Child constraint of rows 3–4.", related=True)
    r = add_param(ws, r, 6, "CaMgtCfg", "Forbid4ccUeRatioThld", "100", "70 (doc optional MML)",
                  "UE-proportion to forbid 4CC or higher.", "Must sit between Forbid3 and Forbid5.", related=True)
    r = add_param(ws, r, 7, "CqiAdaptiveCfg", "SimulAckNackAndCqiFmt3Sw", "See Parameter Reference", "ON if format-3 multiplex wanted",
                  "ACK/NACK + CQI for four CCs.", "")

    r = section(ws, r, COLS, "F.  MML commands  (Ch.7.4.1.2)")
    r = add_mml_header(ws, r)
    for rec in [
        (1, "0", "Both", "FDD/TDD", "Prerequisite",
         "Complete Step1 MML and Step2 CaDl3CCSwitch-1. Group: ≥4 cells. Adaptive: ≥3 SCCs per PCC.",
         "Ch.7.4.1.2"),
        (2, "1", "Both", "FDD/TDD", "Activation",
         'MOD CAMGTCFG: LocalCellId=0, CellCaAlgoSwitch=CaDl4CCSwitch-1;',
         "Repeat every PCell."),
        (3, "2", "Both", "FDD/TDD", "Optional — RLC timers",
         'MOD RLCPDCPPARAGROUP: RlcPdcpParaGroupId=0, RlcMode=RlcMode_AM, Dl4cc5ccUeReorderingTimer=Treordering_m15, Dl4cc5ccUeStatProhTimer=m15;',
         "Doc example values."),
        (4, "9", "Both", "FDD/TDD", "Deactivation",
         'MOD CAMGTCFG: LocalCellId=0, CellCaAlgoSwitch=CaDl4CCSwitch-0;',
         "Leaves 2CC/3CC running."),
    ]:
        r = add_mml(ws, r, *rec)

    r = section(ws, r, COLS, "G.  Counter list  (Tables 7-6, 7-7 + Step1/2)")
    r = add_ctr_header(ws, r)
    for rec in [
        (1, "1526737780", "L.Traffic.User.PCell.DL.4CC.Avg", "Avg UEs in DL 4CC", "Activation verify"),
        (2, "1526737781", "L.Traffic.User.PCell.DL.4CC.Max", "Max UEs in DL 4CC", "Activation"),
        (3, "1526737793", "L.Traffic.User.CA.4CC.PCell.DL.Active.Avg", "Avg UEs with 4 carriers active", "MULTI_CC_STAT_OPT_SW stepping"),
        (4, "1526737794", "L.Traffic.User.CA.4CC.PCell.DL.Active.Max", "Max UEs with 4 carriers active", "Same"),
        (5, "1526737812", "L.Thrp.Time.DL.4CC.CAUser", "DL duration of 4CC UEs", "Throughput denominator"),
        (6, "1526737813", "L.Thrp.bits.DL.4CC.CAUser", "DL bits of 4CC UEs", "Throughput numerator"),
        (7, "—", "(Step1 + Step2 counters)", "Add/act/HO/drop + 3CC counters", "Ch.7.4.3 “in addition to 5.4.3”"),
    ]:
        r = add_ctr(ws, r, *rec)

    r = section(ws, r, COLS, "H.  KPI formulas for 4CC")
    r = _kpi_header(ws, r)
    for i, rec in enumerate([
        ("1", "Throughput of UEs in DL 4CC state",
         "L.Thrp.bits.DL.4CC.CAUser / L.Thrp.Time.DL.4CC.CAUser", "bit/s",
         "Ch.7.4.3. MULTI_CC_STAT_OPT_SW ON: 4 carriers active is enough (5th configured-but-inactive SCell still counts as 4CC)."),
        ("2", "Reuse Step1 KPIs", "Drop / HO / add / act success", "%", "Same formulas"),
    ]):
        vals = list(rec) + [""] * 5
        r = table_row(ws, r, vals, fills=[alt_fill(i)] * 10, height=36)
        merge(ws, r - 1, 5, r - 1, COLS)

    r = section(ws, r, COLS, "I.  Licenses — FDD and TDD  (Ch.7.3.1 / 7.3.2)")
    r = _lic_header(ws, r)
    for i, rec in enumerate([
        ("1", "FDD", "LEOFD-110303", "Carrier Aggregation for Downlink 4CC and 5CC", "LT1SCAD4A5CC", "per cell",
         "Adaptive: PCC-freq cell with CaDl4CCSwitch ON + CA cells on its SCC freqs. Group: ≥4 inter-freq FDD cells and CaDl4CCSwitch on any cell → every group cell. Inter-eNB SCells consume the license on the neighbor eNB. THIS SAME LICENSE ALSO COVERS 5CC (Step4 FDD license = None extra)."),
        ("2", "TDD", "TDLEOFD-081504", "Carrier Aggregation for Downlink 4CC and 5CC", "LT1SCAD4CC00", "per cell",
         "Every TDD cell involved in DL 4CC. Also covers TDD 5CC (Step4 TDD license = None extra)."),
        ("3", "Both", "(Step1+Step2)", "LTE-A Introduction, 2CC 40MHz, 3CC 40/60MHz", "See previous sheets", "per cell",
         "Still required."),
    ]):
        vals = list(rec) + [""] * 3
        r = table_row(ws, r, vals, fills=[alt_fill(i)] * 10, height=52)
        merge(ws, r - 1, 7, r - 1, COLS)

    r = section(ws, r, COLS, "J.  Hardware / networking / verification")
    r = bullets(ws, r, COLS, [
        "Hardware/networking as Step1. UE Rel-12+. Verify Table 7-6 ≠ 0. Tracing: sCellToAddModList with three SCells.",
    ])
    return ws


def build_5cc(wb):
    ws = wb.create_sheet("6. Step4 5CC Config")
    setup_sheet(ws, "5CC")
    set_widths(ws, [8, 28, 32, 22, 28, 55, 42, 12, 12, 12])
    r = 1
    r = banner(ws, r, COLS, "  Step 4 —  Downlink 5CC configuration   (document Ch.8)   —  same structure as Step1")
    r = note_bar(ws, r, COLS, "Prerequisites: 2CC + 3CC (CaDl3CCSwitch) + 4CC (CaDl4CCSwitch). Then CaDl5CCSwitch. "
                 "Group: ≥5 cells. Adaptive: ≥4 candidate SCCs per PCC. "
                 "Ch.8.3.1/8.3.2 Licenses = None extra — FDD/TDD 5CC is covered by the 4CC-and-5CC license already consumed in Step3.")

    r = section(ws, r, COLS, "A.  Principal  (Ch.8.1–8.2)")
    r = insert_figure(ws, r, _fig("emb_p252_0.png"), COLS, "Figure 8-1  Downlink 5CC aggregation  (source: Huawei FPD p.239)")
    r = body(ws, r, COLS,
             "Aggregates five intra- or inter-band carriers. Controlled by CaDl5CCSwitch.\n"
             "PUCCH format-3 multiplex: SimulAckNackAndCqiFmt3Sw. Overhead/IBLER: same as 3CC/4CC.\n"
             "Hardware extra vs 2CC: among BBPs, UBBP is recommended; among main control boards, UMPT is recommended (Ch.8.3.4).")
    r = peak_chart(ws, r, "Theoretical peak data rates for DL 5CC  (Table 8-1, Mbit/s)",
                   ["2x2 + 64QAM", "2x2 + 256QAM", "4x4 + 64QAM", "4x4 + 256QAM"],
                   [748.9, 979.1, 1499.3, 1958.3],
                   [560, 715, 1085, 1430], "5cc")

    r = trigger_leave_block(ws, r, "5CC")

    r = section(ws, r, COLS, "D.  Prerequisite, mutually exclusive, and mutually impacted features")
    r = headers(ws, r, ["SN", "Relation", "RAT", "Feature", "Parameter(s)", "Required action", "Impact / note"] + [""] * 3)
    for rec in [
        ("1", "Prerequisite", "FDD/TDD", "Downlink 2CC", "None extra", "Step1 ON", "Ch.8.3.3.1"),
        ("2", "Prerequisite", "FDD/TDD", "Downlink 3CC", "CaDl3CCSwitch", "Step2 ON", "Same MO as CaDl5CCSwitch", True),
        ("3", "Prerequisite", "FDD/TDD", "Downlink 4CC", "CaDl4CCSwitch", "Step3 ON", "Same MO", True),
        ("4", "Exclusive", "FDD", "PUCCH measurement", "PucchMeasOptSwitch OFF", "Deactivate", "Ch.8.3.3.2"),
        ("5", "Impact", "FDD", "UMTS/LTE Zero Bufferzone", "UMTS_LTE_ZERO_BUFFER_ZONE_SW", "Avoid 5/10 MHz bufferzone PCell", ""),
        ("6", "Impact", "FDD", "TDM (NSA performance enhancement)", "NsaDcMgmtConfig.NsaDcAlgoSwitch → TDM_SWITCH",
         "Max CC count vs TDD UL/DL config: cfg 0/1/6 → max 5; cfg 2 → max 4; cfg 3/4 → max 3; cfg 5 → max 2. Exceeding → SCells removed if TDM required",
         "5CC-specific. If TDM ON with UL/DL cfg 2, 5CC cannot stay at 5."),
        ("7", "Impact", "FDD", "UMTS/LTE Spectrum Sharing", "UL_SPECTRUM_SHARING", "5 MHz not recommended as PCell", ""),
        ("8", "Impact", "FDD", "UMTS/LTE SS based on DC-HSDPA", "DC_HSDPA_BASED_UL_SPECTRUM_SHR", "5 MHz not recommended as PCell", "Related", True),
        ("9", "Impact", "TDD", "SRS resource migration", "REMOTE_INTRF_BF_ENH_SW", "SRS migration pre-empts 5CC", ""),
        ("10", "Constraint", "FDD/TDD", "UE / EPC", "UE Rel-12+; MBR ≥ Table 8-1", "", ""),
        ("11", "Inherit", "FDD/TDD", "Step1 exclusives", "Super combined cell, >100 km, in-band relay, MBSFN, MIE ES", "Keep OFF", ""),
    ]:
        related = rec[-1] is True
        rec = rec[:-1] if rec[-1] is True else rec
        vals = list(rec) + [""] * 3
        rel = rec[1]
        fh = PALE_GREEN if rel == "Prerequisite" else (PALE_RED if rel == "Exclusive" else (PALE_GOLD if related or rel in ("Constraint", "Inherit") else alt_fill(int(rec[0]))))
        r = table_row(ws, r, vals, fills=[fh] * 10, height=40)
        merge(ws, r - 1, 7, r - 1, COLS)

    r = section(ws, r, COLS, "E.  All parameters  (5CC-specific on top of Steps 1–3)")
    r = add_param_header(ws, r)
    r = add_param(ws, r, 1, "CaMgtCfg", "CellCaAlgoSwitch / CaDl5CCSwitch", "OFF", "1 on every possible PCell",
                  "Master switch for downlink 5CC.", "Requires CaDl3CCSwitch and CaDl4CCSwitch already 1.")
    r = add_param(ws, r, 2, "CaMgtCfg", "CellCaAlgoSwitch / CaDl4CCSwitch", "1 after Step3", "Keep 1",
                  "4CC prerequisite bit.", "Related same MO.", related=True)
    r = add_param(ws, r, 3, "CaMgtCfg", "CellCaAlgoSwitch / CaDl3CCSwitch", "1 after Step2", "Keep 1",
                  "3CC prerequisite bit.", "Related same MO.", related=True)
    r = add_param(ws, r, 4, "RlcPdcpParaGroup", "Dl4cc5ccUeReorderingTimer", "See Parameter Reference", "Treordering_m15 (Table 8-3)",
                  "Same beyond-3CC reordering timer as 4CC.", "Requires RlcMode_AM.", related=True)
    r = add_param(ws, r, 5, "RlcPdcpParaGroup", "Dl4cc5ccUeStatProhTimer", "See Parameter Reference", "m15 (Table 8-3)",
                  "Same status-prohibit timer as 4CC.", "Related to row 4.", related=True)
    r = add_param(ws, r, 6, "CaMgtCfg", "Forbid5ccUeRatioThld", "100", "60 (doc optional MML)",
                  "UE-proportion to forbid 5CC or higher.", "Lowest of the Forbid* ladder.", related=True)
    r = add_param(ws, r, 7, "NsaDcMgmtConfig", "NsaDcAlgoSwitch / TDM_SWITCH", "OFF unless NSA TDM is planned", "See impact row 6; may cap CC count below 5",
                  "TDM ACK multiplexing constraint vs UL/DL configuration.", "Related 5CC-specific impact.", related=True)
    r = add_param(ws, r, 8, "CqiAdaptiveCfg", "SimulAckNackAndCqiFmt3Sw", "See Parameter Reference", "ON if format-3 multiplex wanted",
                  "ACK/NACK + CQI for five CCs.", "")

    r = section(ws, r, COLS, "F.  MML commands  (Ch.8.4.1.2)")
    r = add_mml_header(ws, r)
    for rec in [
        (1, "0", "Both", "FDD/TDD", "Prerequisite",
         "Complete Step1–3. Group: ≥5 cells. Adaptive: ≥4 SCCs per PCC. UBBP + UMPT recommended.",
         "Ch.8.4.1.2"),
        (2, "1", "Both", "FDD/TDD", "Activation",
         'MOD CAMGTCFG: LocalCellId=0, CellCaAlgoSwitch=CaDl5CCSwitch-1;',
         "Repeat every PCell."),
        (3, "2", "Both", "FDD/TDD", "Optional — RLC timers",
         'MOD RLCPDCPPARAGROUP: RlcPdcpParaGroupId=0, RlcMode=RlcMode_AM, Dl4cc5ccUeReorderingTimer=Treordering_m15, Dl4cc5ccUeStatProhTimer=m15;',
         "Same command as 4CC (shared timers)."),
        (4, "9", "Both", "FDD/TDD", "Deactivation",
         'MOD CAMGTCFG: LocalCellId=0, CellCaAlgoSwitch=CaDl5CCSwitch-0;',
         "Leaves 2/3/4CC running."),
    ]:
        r = add_mml(ws, r, *rec)

    r = section(ws, r, COLS, "G.  Counter list  (Tables 8-4, 8-5 + previous steps)")
    r = add_ctr_header(ws, r)
    for rec in [
        (1, "1526739805", "L.Traffic.User.PCell.DL.5CC.Avg", "Avg UEs in DL 5CC", "Activation verify"),
        (2, "1526739806", "L.Traffic.User.PCell.DL.5CC.Max", "Max UEs in DL 5CC", "Activation"),
        (3, "1526740442", "L.Traffic.User.CA.5CC.PCell.DL.Active.Avg", "Avg UEs with 5 carriers active", "MULTI_CC_STAT_OPT_SW stepping"),
        (4, "1526740443", "L.Traffic.User.CA.5CC.PCell.DL.Active.Max", "Max UEs with 5 carriers active", "Same"),
        (5, "1526740438", "L.Thrp.Time.DL.5CC.CAUser", "DL duration of 5CC UEs", "Throughput denominator"),
        (6, "1526740439", "L.Thrp.bits.DL.5CC.CAUser", "DL bits of 5CC UEs", "Throughput numerator"),
        (7, "—", "(Step1–3 counters)", "Add/act/HO/drop + 3CC + 4CC", "Ch.8.4.3 “in addition to 5.4.3”"),
    ]:
        r = add_ctr(ws, r, *rec)

    r = section(ws, r, COLS, "H.  KPI formulas for 5CC")
    r = _kpi_header(ws, r)
    for i, rec in enumerate([
        ("1", "Throughput of UEs in DL 5CC state",
         "L.Thrp.bits.DL.5CC.CAUser / L.Thrp.Time.DL.5CC.CAUser", "bit/s",
         "Ch.8.4.3. MULTI_CC_STAT_OPT_SW ON: counted when 5 carriers are active (a 6th inactive configured SCell still counts as 5CC)."),
        ("2", "Reuse Step1 KPIs", "Drop / HO / add / act success", "%", "Same formulas"),
    ]):
        vals = list(rec) + [""] * 5
        r = table_row(ws, r, vals, fills=[alt_fill(i)] * 10, height=36)
        merge(ws, r - 1, 5, r - 1, COLS)

    r = section(ws, r, COLS, "I.  Licenses — FDD and TDD  (Ch.8.3.1 / 8.3.2)")
    r = _lic_header(ws, r)
    for i, rec in enumerate([
        ("1", "FDD", "—", "None additional", "—", "—",
         "Ch.8.3.1 Licenses (FDD) = None. Covered by LEOFD-110303 (LT1SCAD4A5CC) already required in Step3."),
        ("2", "TDD", "—", "None additional", "—", "—",
         "Ch.8.3.2 Licenses (TDD) = None. Covered by TDLEOFD-081504 (LT1SCAD4CC00) already required in Step3."),
        ("3", "Both", "(Steps 1–3)", "LTE-A Introduction, 2CC 40MHz, 3CC, 4CC+5CC", "See previous sheets", "per cell",
         "All previous licenses remain consumed."),
    ]):
        vals = list(rec) + [""] * 3
        r = table_row(ws, r, vals, fills=[alt_fill(i)] * 10, height=40)
        merge(ws, r - 1, 7, r - 1, COLS)

    r = section(ws, r, COLS, "J.  Hardware / networking / verification")
    r = bullets(ws, r, COLS, [
        "Same networking as Step1. Extra: UBBP recommended, UMPT recommended.",
        "UE Rel-12+. Verify Table 8-4 ≠ 0. Tracing: sCellToAddModList with four SCells.",
        "Watch TDM_SWITCH vs UL/DL configuration if NSA TDM is on — 5CC may be forced down to 4/3/2 CCs.",
    ])
    return ws
