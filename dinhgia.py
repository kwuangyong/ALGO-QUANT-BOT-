"""
valuation_cfa.py — CFA L2-grade intrinsic valuation module
============================================================
Input  : Excel file từ quant_pipeline.py (cột 'ticker', 'signal')
Output : Word report tiếng Việt (5-paragraph template per ticker)

Methods (confidence-weighted triangulation):
    1. FCFF 2-stage DCF        (industrial, consumer, tech, utilities)
    2. Residual Income (EBO)   (banking, insurance, BV-heavy)
    3. Relative valuation      (P/E, P/B, EV/EBITDA — sanity check)

Author: Naizy Quant Pipeline — Valuation Layer
"""

from __future__ import annotations

import logging
import warnings
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path
from typing import Optional

import numpy as np
import pandas as pd

warnings.filterwarnings("ignore")
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(name)s: %(message)s",
    datefmt="%H:%M:%S",
)
log = logging.getLogger("valuation_cfa")

# ============================================================
# CONFIG
# ============================================================

CFG = {
    # --- Input ---
    "ticker_col": "ticker",
    "signal_col": "signal",
    "buy_values": ["BUY", "STRONG_BUY", "MUA", "MUA_MANH"],

    # --- Macro VN (Damodaran 2026 + SBV) ---
    "rf":           0.045,    # TPCP 10Y
    "erp_vn":       0.085,    # Frontier market ERP
    "g_terminal":   0.040,    # Inflation dài hạn VN
    "tax_rate":     0.20,     # CIT VN
    "default_beta": 1.0,
    "high_growth_years": 5,   # Stage 1 DCF horizon

    # --- Method weights (BASE — sẽ adjust theo confidence) ---
    "method_weights_base": {
        "Banking":      {"RI": 0.60, "PB": 0.30, "PE": 0.10, "FCFF": 0.00},
        "Insurance":    {"RI": 0.55, "PB": 0.30, "PE": 0.15, "FCFF": 0.00},
        "Real_Estate":  {"FCFF": 0.40, "PB": 0.35, "RI": 0.25, "PE": 0.00},
        "Securities":   {"PB": 0.45, "RI": 0.35, "PE": 0.20, "FCFF": 0.00},
        "Industrial":   {"FCFF": 0.50, "PE": 0.30, "RI": 0.20, "PB": 0.00},
        "Consumer":     {"FCFF": 0.45, "PE": 0.35, "RI": 0.20, "PB": 0.00},
        "Utilities":    {"FCFF": 0.40, "PE": 0.30, "RI": 0.20, "PB": 0.10},
        "Tech":         {"FCFF": 0.55, "PE": 0.30, "RI": 0.15, "PB": 0.00},
        "Materials":    {"FCFF": 0.45, "PE": 0.30, "RI": 0.20, "PB": 0.05},
        "Healthcare":   {"FCFF": 0.50, "PE": 0.30, "RI": 0.20, "PB": 0.00},
        "default":      {"FCFF": 0.40, "PE": 0.30, "RI": 0.20, "PB": 0.10},
    },

    # --- Sector P/E & P/B benchmarks VN (refresh quarterly) ---
    "sector_multiples": {
        "Banking":      {"PE": 9.0,  "PB": 1.6, "EV_EBITDA": None},
        "Insurance":    {"PE": 12.0, "PB": 1.8, "EV_EBITDA": None},
        "Real_Estate":  {"PE": 14.0, "PB": 1.5, "EV_EBITDA": 12.0},
        "Securities":   {"PE": 13.0, "PB": 1.7, "EV_EBITDA": None},
        "Industrial":   {"PE": 12.0, "PB": 1.8, "EV_EBITDA": 8.0},
        "Consumer":     {"PE": 16.0, "PB": 2.5, "EV_EBITDA": 10.0},
        "Utilities":    {"PE": 11.0, "PB": 1.4, "EV_EBITDA": 7.0},
        "Tech":         {"PE": 20.0, "PB": 3.5, "EV_EBITDA": 13.0},
        "Materials":    {"PE": 10.0, "PB": 1.5, "EV_EBITDA": 7.5},
        "Healthcare":   {"PE": 18.0, "PB": 2.8, "EV_EBITDA": 12.0},
        "default":      {"PE": 13.0, "PB": 1.8, "EV_EBITDA": 9.0},
    },

    # --- Quality gates (flag, không drop) ---
    "quality_gates": {
        "min_roe_3y":         0.05,
        "max_debt_ebitda":    5.0,
        "min_interest_cover": 2.0,
        "max_neg_fcf_years":  2,
    },

    # --- Verdict thresholds ---
    "mos_attractive": 0.20,    # MOS > +20% → 🟢
    "mos_expensive":  -0.10,   # MOS ≤ -10% → 🔴

    # --- Sensitivity grid ---
    "sens_g":    [0.02, 0.03, 0.04, 0.05],
    "sens_wacc": [-0.01, -0.005, 0.0, 0.005, 0.01],

    # --- Output ---
    "output_dir": "./output",
}


# ============================================================
# DATA STRUCTURES
# ============================================================

@dataclass
class FinancialData:
    """Container cho dữ liệu BCTC của 1 mã (đã normalize)."""
    ticker: str
    sector: str

    # Market data
    market_price:     float = np.nan
    shares_out:       float = np.nan        # million shares
    market_cap:       float = np.nan        # billion VND

    # Income statement (3-5Y history)
    revenue:          list[float] = field(default_factory=list)
    ebit:             list[float] = field(default_factory=list)
    net_income:       list[float] = field(default_factory=list)
    interest_expense: list[float] = field(default_factory=list)
    da:               list[float] = field(default_factory=list)  # depreciation+amortization

    # Balance sheet
    total_assets:     list[float] = field(default_factory=list)
    total_equity:     list[float] = field(default_factory=list)
    total_debt:       list[float] = field(default_factory=list)
    working_capital:  list[float] = field(default_factory=list)
    book_value_ps:    float = np.nan
    eps_ttm:          float = np.nan

    # Cash flow
    capex:            list[float] = field(default_factory=list)
    fcf:              list[float] = field(default_factory=list)

    # Beta (rolling 2Y vs VN-Index)
    beta:             float = np.nan

    # Quality flags
    quality_flags:    list[str] = field(default_factory=list)

    def has_minimum_data(self) -> tuple[bool, str]:
        """Kiểm tra data tối thiểu — KHÔNG drop, chỉ log."""
        if np.isnan(self.market_price):
            return False, "market_price missing"
        if len(self.net_income) < 3:
            return False, f"net_income chỉ có {len(self.net_income)} năm (cần ≥3)"
        if np.isnan(self.book_value_ps) or self.book_value_ps <= 0:
            return False, "book_value_ps invalid"
        return True, "ok"


@dataclass
class MethodResult:
    """Kết quả 1 method định giá."""
    method:     str
    fair_value: float
    confidence: float       # 0.0–1.0, dùng cho weighted avg
    details:    dict = field(default_factory=dict)
    error:      Optional[str] = None

    @property
    def success(self) -> bool:
        return self.error is None and not np.isnan(self.fair_value) and self.fair_value > 0


@dataclass
class ValuationResult:
    """Kết quả tổng hợp cho 1 ticker."""
    ticker:           str
    sector:           str
    market_price:     float
    fair_value:       float = np.nan
    mos:              float = np.nan      # margin of safety
    verdict:          str = "N/A"          # 🟢 / 🟡 / 🔴
    methods:          dict[str, MethodResult] = field(default_factory=dict)
    weights_applied:  dict[str, float] = field(default_factory=dict)
    quality_flags:    list[str] = field(default_factory=list)
    cost_of_equity:   float = np.nan
    beta:             float = np.nan
    sensitivity:      Optional[pd.DataFrame] = None
    status:           str = "PENDING"      # SUCCESS | PARTIAL | FAILED
    fail_reason:      Optional[str] = None
    # Quant signals từ buy_summary.xlsx (Score, Rating, HMM, Forecast, ...)
    quant_signals:    dict = field(default_factory=dict)


# ============================================================
# SECTOR CLASSIFICATION
# ============================================================

# HOSE/HNX ticker → sector map (extended for buy_summary tickers)
SECTOR_MAP = {
    # Banking
    "VCB": "Banking", "BID": "Banking", "CTG": "Banking", "TCB": "Banking",
    "MBB": "Banking", "ACB": "Banking", "VPB": "Banking", "HDB": "Banking",
    "STB": "Banking", "TPB": "Banking", "VIB": "Banking", "SHB": "Banking",
    "LPB": "Banking", "OCB": "Banking", "MSB": "Banking", "EIB": "Banking",
    "NAB": "Banking", "BAB": "Banking", "ABB": "Banking", "KLB": "Banking",
    "NVB": "Banking", "SGB": "Banking", "VBB": "Banking", "PGB": "Banking",
    # Insurance
    "BVH": "Insurance", "BMI": "Insurance", "MIG": "Insurance", "PVI": "Insurance",
    "PTI": "Insurance", "VNR": "Insurance", "ABI": "Insurance",
    # Real Estate
    "VHM": "Real_Estate", "VIC": "Real_Estate", "VRE": "Real_Estate",
    "NVL": "Real_Estate", "KDH": "Real_Estate", "DXG": "Real_Estate",
    "PDR": "Real_Estate", "NLG": "Real_Estate", "DIG": "Real_Estate",
    "HDG": "Real_Estate", "KBC": "Real_Estate", "BCM": "Real_Estate",
    "VPI": "Real_Estate", "VPL": "Real_Estate", "LGL": "Real_Estate",
    "SIP": "Real_Estate", "IDC": "Real_Estate", "SZC": "Real_Estate",
    "TIG": "Real_Estate", "CEO": "Real_Estate", "HQC": "Real_Estate",
    # Securities
    "SSI": "Securities", "VND": "Securities", "VCI": "Securities", "HCM": "Securities",
    "SHS": "Securities", "VIX": "Securities", "MBS": "Securities", "FTS": "Securities",
    "BSI": "Securities", "AGR": "Securities", "CTS": "Securities", "ORS": "Securities",
    # Consumer
    "VNM": "Consumer", "MSN": "Consumer", "SAB": "Consumer", "MWG": "Consumer",
    "PNJ": "Consumer", "DGW": "Consumer", "FRT": "Consumer", "VEA": "Consumer",
    "TLG": "Consumer", "QNS": "Consumer", "KDC": "Consumer", "VHC": "Consumer",
    "ANV": "Consumer", "IDI": "Consumer", "ASM": "Consumer", "NAF": "Consumer",
    "HAG": "Consumer", "HNG": "Consumer",
    # Tech
    "FPT": "Tech", "CMG": "Tech", "ELC": "Tech", "ITD": "Tech",
    # Utilities
    "GAS": "Utilities", "POW": "Utilities", "PGV": "Utilities", "NT2": "Utilities",
    "REE": "Utilities", "PC1": "Utilities", "GEX": "Utilities", "GEG": "Utilities",
    "PPC": "Utilities", "VSH": "Utilities", "HND": "Utilities",
    # Materials / Oil & Gas
    "HPG": "Materials", "HSG": "Materials", "NKG": "Materials", "DCM": "Materials",
    "DPM": "Materials", "DGC": "Materials", "BMP": "Materials", "BSR": "Materials",
    "PVS": "Materials", "PVD": "Materials", "PVT": "Materials", "PLX": "Materials",
    "OIL": "Materials", "PVB": "Materials",
    # Industrial
    "GMD": "Industrial", "VSC": "Industrial", "HAH": "Industrial", "VTP": "Industrial",
    "ACV": "Industrial", "HVN": "Industrial", "VJC": "Industrial",
    "HHP": "Industrial", "HII": "Industrial", "PTB": "Industrial",
    "TSA": "Industrial", "DHC": "Industrial", "VGS": "Industrial",
    "TNT": "Industrial", "GEL": "Industrial", "VCS": "Industrial",
    # Healthcare / Pharma
    "DHG": "Healthcare", "IMP": "Healthcare", "DBD": "Healthcare", "TRA": "Healthcare",
    "DCL": "Healthcare", "PME": "Healthcare", "OPC": "Healthcare", "DMC": "Healthcare",
}


def classify_sector(ticker: str) -> str:
    """Trả về sector của ticker. Fallback 'default' nếu không map được."""
    sector = SECTOR_MAP.get(ticker.upper(), "default")
    if sector == "default":
        log.warning(f"[{ticker}] không có sector mapping — dùng 'default' weights")
    return sector


# ============================================================
# DATA LOADER (vnstock golden tier)
# ============================================================

def load_financial_data(ticker: str) -> FinancialData:
    """
    Load BCTC + market data từ vnstock golden tier.
    Fail-loud: raise rõ ràng nếu thiếu data critical.
    """
    sector = classify_sector(ticker)
    fd = FinancialData(ticker=ticker, sector=sector)

    try:
        from vnstock import Vnstock
        v = Vnstock().stock(symbol=ticker, source="VCI")

        # --- Market price (latest close) ---
        try:
            quote = v.quote.history(
                start=(datetime.now() - pd.Timedelta(days=10)).strftime("%Y-%m-%d"),
                end=datetime.now().strftime("%Y-%m-%d"),
            )
            if quote is not None and len(quote) > 0:
                fd.market_price = float(quote["close"].iloc[-1])
        except Exception as e:
            log.warning(f"[{ticker}] market price fetch failed: {e}")

        # --- Income Statement ---
        try:
            is_df = v.finance.income_statement(period="year", lang="en")
            if is_df is not None and len(is_df) > 0:
                # vnstock returns most-recent-first; reverse for chronological
                is_df = is_df.sort_index() if "yearReport" not in is_df.columns \
                    else is_df.sort_values("yearReport")
                fd.revenue          = _safe_col(is_df, ["Revenue (Bn. VND)", "Revenue", "revenue"])
                fd.net_income       = _safe_col(is_df, ["Attribute to parent company (Bn. VND)",
                                                         "Net Profit For the Year", "net_income"])
                fd.ebit             = _safe_col(is_df, ["Profit before tax", "EBIT", "ebit"])
                fd.interest_expense = _safe_col(is_df, ["Interest Expenses", "interest_expense"])
        except Exception as e:
            log.warning(f"[{ticker}] income statement failed: {e}")

        # --- Balance Sheet ---
        try:
            bs_df = v.finance.balance_sheet(period="year", lang="en")
            if bs_df is not None and len(bs_df) > 0:
                bs_df = bs_df.sort_index() if "yearReport" not in bs_df.columns \
                    else bs_df.sort_values("yearReport")
                fd.total_assets    = _safe_col(bs_df, ["TOTAL ASSETS (Bn. VND)", "total_assets"])
                fd.total_equity    = _safe_col(bs_df, ["OWNER'S EQUITY(Bn.VND)", "total_equity"])
                fd.total_debt      = _safe_col(bs_df, ["Long-term borrowings (Bn. VND)",
                                                        "Short-term borrowings (Bn. VND)",
                                                        "total_debt"], aggregate="sum")
        except Exception as e:
            log.warning(f"[{ticker}] balance sheet failed: {e}")

        # --- Cash Flow ---
        try:
            cf_df = v.finance.cash_flow(period="year", lang="en")
            if cf_df is not None and len(cf_df) > 0:
                cf_df = cf_df.sort_index() if "yearReport" not in cf_df.columns \
                    else cf_df.sort_values("yearReport")
                fd.da    = _safe_col(cf_df, ["Depreciation and Amortisation", "da"])
                fd.capex = _safe_col(cf_df, ["Purchase of fixed assets", "capex"])
        except Exception as e:
            log.warning(f"[{ticker}] cash flow failed: {e}")

        # --- Ratios (EPS, BVPS) ---
        try:
            ratio_df = v.finance.ratio(period="year", lang="en")
            if ratio_df is not None and len(ratio_df) > 0:
                ratio_df = ratio_df.sort_index() if "yearReport" not in ratio_df.columns \
                    else ratio_df.sort_values("yearReport")
                eps_series = _safe_col(ratio_df, ["EPS (VND)", "eps"])
                bv_series  = _safe_col(ratio_df, ["BVPS (VND)", "book_value_per_share"])
                if eps_series: fd.eps_ttm = eps_series[-1]
                if bv_series:  fd.book_value_ps = bv_series[-1]
        except Exception as e:
            log.warning(f"[{ticker}] ratios failed: {e}")

        # --- Compute FCF from FCFF formula nếu chưa có ---
        if not fd.fcf and fd.net_income and fd.da and fd.capex:
            n = min(len(fd.net_income), len(fd.da), len(fd.capex))
            for i in range(n):
                fcf_i = fd.net_income[i] + fd.da[i] - abs(fd.capex[i])
                fd.fcf.append(fcf_i)

    except ImportError:
        log.error("vnstock chưa cài. pip install vnstock")
        raise
    except Exception as e:
        log.error(f"[{ticker}] data loader hard fail: {e}", exc_info=True)

    return fd


def _safe_col(df: pd.DataFrame, candidates: list[str], aggregate: str = None) -> list[float]:
    """Tìm cột theo danh sách tên candidate, trả về list[float] đã clean NaN."""
    found_cols = []
    for col in candidates:
        if col in df.columns:
            found_cols.append(col)
    if not found_cols:
        return []
    if aggregate == "sum" and len(found_cols) > 1:
        series = df[found_cols].sum(axis=1)
    else:
        series = df[found_cols[0]]
    return [float(x) for x in series.dropna().tolist() if not pd.isna(x)]


# ============================================================
# COST OF EQUITY (CAPM)
# ============================================================

def compute_cost_of_equity(fd: FinancialData) -> tuple[float, float]:
    """
    CAPM: r_e = rf + β × ERP_VN
    Returns: (cost_of_equity, beta_used)
    """
    beta = fd.beta if not np.isnan(fd.beta) else CFG["default_beta"]
    re = CFG["rf"] + beta * CFG["erp_vn"]
    return re, beta


def compute_wacc(fd: FinancialData, re: float) -> float:
    """WACC = (E/V)×re + (D/V)×rd×(1-t)"""
    if not fd.total_debt or not fd.total_equity:
        return re  # all-equity assumption
    D = fd.total_debt[-1]
    E = fd.total_equity[-1] if fd.total_equity[-1] > 0 else 1.0
    V = D + E
    if V <= 0:
        return re
    # rd proxy: interest expense / avg debt (fallback 8%)
    rd = 0.08
    if fd.interest_expense and len(fd.total_debt) >= 2:
        avg_debt = (fd.total_debt[-1] + fd.total_debt[-2]) / 2
        if avg_debt > 0:
            rd = fd.interest_expense[-1] / avg_debt
            rd = max(0.04, min(rd, 0.15))  # clip 4-15%
    return (E/V)*re + (D/V)*rd*(1 - CFG["tax_rate"])


# ============================================================
# METHOD 1: FCFF 2-STAGE DCF
# ============================================================

def valuate_fcff(fd: FinancialData, wacc: float) -> MethodResult:
    """FCFF 2-stage DCF. Skip cho banking/insurance."""
    if fd.sector in ("Banking", "Insurance", "Securities"):
        return MethodResult("FCFF", np.nan, 0.0, error="N/A cho financial sector")

    if not fd.fcf or len(fd.fcf) < 3:
        return MethodResult("FCFF", np.nan, 0.0,
                            error=f"FCF chỉ có {len(fd.fcf)} năm (cần ≥3)")

    try:
        # --- Growth rate stage 1: 3Y CAGR, clip [-5%, +25%] ---
        fcf_recent = fd.fcf[-3:]
        if fcf_recent[0] <= 0:
            # FCF âm: dùng revenue growth làm proxy
            if len(fd.revenue) >= 3 and fd.revenue[-3] > 0:
                g1 = (fd.revenue[-1] / fd.revenue[-3]) ** (1/2) - 1
            else:
                g1 = 0.05
        else:
            g1 = (fcf_recent[-1] / fcf_recent[0]) ** (1/2) - 1
        g1 = float(np.clip(g1, -0.05, 0.25))

        g_terminal = CFG["g_terminal"]
        if g_terminal >= wacc:
            return MethodResult("FCFF", np.nan, 0.0,
                                error=f"g_terminal ({g_terminal:.1%}) ≥ WACC ({wacc:.1%})")

        # --- Project FCF stage 1 ---
        fcf0 = abs(fd.fcf[-1]) if fd.fcf[-1] > 0 else max(abs(np.mean(fd.fcf[-3:])), 1e-6)
        proj_fcf = []
        for t in range(1, CFG["high_growth_years"] + 1):
            # Linear fade từ g1 về g_terminal
            g_t = g1 + (g_terminal - g1) * (t / CFG["high_growth_years"])
            fcf_t = fcf0 * np.prod([1 + g1 + (g_terminal - g1) * (i / CFG["high_growth_years"])
                                     for i in range(1, t+1)])
            proj_fcf.append(fcf_t)

        # --- PV stage 1 ---
        pv_stage1 = sum(fcf / (1 + wacc)**t for t, fcf in enumerate(proj_fcf, start=1))

        # --- Terminal value ---
        tv = proj_fcf[-1] * (1 + g_terminal) / (wacc - g_terminal)
        pv_tv = tv / (1 + wacc)**CFG["high_growth_years"]

        # --- Enterprise value → Equity value ---
        ev = pv_stage1 + pv_tv
        net_debt = fd.total_debt[-1] if fd.total_debt else 0
        equity_value = ev - net_debt

        if equity_value <= 0 or np.isnan(fd.shares_out) or fd.shares_out <= 0:
            # Fallback shares_out từ market_cap / price
            if not np.isnan(fd.market_price) and fd.market_price > 0:
                if fd.net_income and fd.eps_ttm and fd.eps_ttm > 0:
                    shares_est = fd.net_income[-1] * 1e9 / fd.eps_ttm  # NI in Bn VND × 1e9 = VND
                    fair_value_ps = equity_value * 1e9 / shares_est
                else:
                    return MethodResult("FCFF", np.nan, 0.0, error="shares_out missing")
            else:
                return MethodResult("FCFF", np.nan, 0.0, error="cannot compute per-share")
        else:
            fair_value_ps = equity_value * 1e9 / (fd.shares_out * 1e6)

        # --- Confidence score ---
        conf = _confidence_fcff(fd, g1, wacc)

        return MethodResult(
            "FCFF", fair_value_ps, conf,
            details={
                "g_stage1": g1, "g_terminal": g_terminal, "wacc": wacc,
                "pv_stage1_bn": pv_stage1, "pv_terminal_bn": pv_tv,
                "equity_value_bn": equity_value, "fcf0_bn": fcf0,
            }
        )
    except Exception as e:
        log.exception(f"[{fd.ticker}] FCFF error: {e}")
        return MethodResult("FCFF", np.nan, 0.0, error=str(e))


def _confidence_fcff(fd: FinancialData, g1: float, wacc: float) -> float:
    """Confidence FCFF: high khi data đầy đủ, FCF dương ổn định."""
    conf = 1.0
    if len(fd.fcf) < 5:           conf *= 0.85
    if sum(1 for x in fd.fcf[-3:] if x < 0) > 0:  conf *= 0.7
    if abs(g1) > 0.20:            conf *= 0.75
    if wacc - CFG["g_terminal"] < 0.03:  conf *= 0.6  # quá close to terminal
    return float(np.clip(conf, 0.1, 1.0))


# ============================================================
# METHOD 2: RESIDUAL INCOME (Edwards-Bell-Ohlson)
# ============================================================

def valuate_ri(fd: FinancialData, re: float) -> MethodResult:
    """V₀ = BV₀ + Σ PV(RI_t). RI_t = (ROE_t - r) × BV_{t-1}"""
    if not fd.total_equity or len(fd.total_equity) < 3:
        return MethodResult("RI", np.nan, 0.0,
                            error=f"equity chỉ có {len(fd.total_equity)} năm")
    if not fd.net_income or len(fd.net_income) < 3:
        return MethodResult("RI", np.nan, 0.0, error="net_income < 3 năm")
    if np.isnan(fd.book_value_ps) or fd.book_value_ps <= 0:
        return MethodResult("RI", np.nan, 0.0, error="BVPS invalid")

    try:
        # --- ROE bền vững = trung bình 3Y ---
        n_overlap = min(len(fd.net_income), len(fd.total_equity))
        roe_series = []
        for i in range(-min(3, n_overlap), 0):
            eq = fd.total_equity[i]
            if eq > 0:
                roe_series.append(fd.net_income[i] / eq)
        if not roe_series:
            return MethodResult("RI", np.nan, 0.0, error="không tính được ROE")
        roe_sustain = float(np.mean(roe_series))

        # --- Edge case: ROE ≤ r → fair value = BV (no excess return) ---
        if roe_sustain <= re:
            fair_value = fd.book_value_ps
            return MethodResult(
                "RI", fair_value, 0.5,
                details={"roe_sustain": roe_sustain, "re": re,
                         "note": "ROE ≤ r → fair = BV"}
            )

        # --- Persistence factor ω = 0.6 (mid-range, Vietnamese market) ---
        omega = 0.6
        # Clean surplus: V₀ = BV₀ + (ROE-r)×BV₀ / (1+r-ω)
        excess_return_perpetuity = (roe_sustain - re) * fd.book_value_ps / (1 + re - omega)
        fair_value = fd.book_value_ps + excess_return_perpetuity

        conf = _confidence_ri(fd, roe_series, re)

        return MethodResult(
            "RI", fair_value, conf,
            details={
                "bv_ps": fd.book_value_ps, "roe_sustain": roe_sustain,
                "re": re, "omega_persistence": omega,
                "excess_return_ps": excess_return_perpetuity,
            }
        )
    except Exception as e:
        log.exception(f"[{fd.ticker}] RI error: {e}")
        return MethodResult("RI", np.nan, 0.0, error=str(e))


def _confidence_ri(fd: FinancialData, roe_series: list[float], re: float) -> float:
    conf = 1.0
    if len(roe_series) < 3:                          conf *= 0.7
    if np.std(roe_series) > 0.05:                    conf *= 0.8   # ROE biến động
    if np.mean(roe_series) - re < 0.02:              conf *= 0.7   # excess return mỏng
    if fd.sector in ("Banking", "Insurance"):        conf *= 1.10  # phương pháp phù hợp
    return float(np.clip(conf, 0.1, 1.0))


# ============================================================
# METHOD 3: RELATIVE VALUATION (P/E, P/B, EV/EBITDA)
# ============================================================

def valuate_relative(fd: FinancialData) -> dict[str, MethodResult]:
    """Trả về dict {PE, PB, EV_EBITDA} — mỗi cái là 1 MethodResult."""
    results = {}
    benchmarks = CFG["sector_multiples"].get(fd.sector, CFG["sector_multiples"]["default"])

    # --- P/E ---
    if not np.isnan(fd.eps_ttm) and fd.eps_ttm > 0:
        pe_target = benchmarks["PE"]
        # Smooth EPS = mean 3Y
        eps_smooth = fd.eps_ttm
        if fd.net_income and len(fd.net_income) >= 3 and not np.isnan(fd.shares_out):
            ni_avg = np.mean(fd.net_income[-3:])
            eps_smooth = (ni_avg * 1e9) / (fd.shares_out * 1e6)
        fair_pe = eps_smooth * pe_target
        conf = _confidence_relative(fd, "PE")
        results["PE"] = MethodResult("PE", fair_pe, conf,
                                      details={"eps_smooth": eps_smooth, "pe_target": pe_target})
    else:
        results["PE"] = MethodResult("PE", np.nan, 0.0, error="EPS invalid hoặc âm")

    # --- P/B ---
    if not np.isnan(fd.book_value_ps) and fd.book_value_ps > 0:
        pb_target = benchmarks["PB"]
        fair_pb = fd.book_value_ps * pb_target
        conf = _confidence_relative(fd, "PB")
        results["PB"] = MethodResult("PB", fair_pb, conf,
                                      details={"bv_ps": fd.book_value_ps, "pb_target": pb_target})
    else:
        results["PB"] = MethodResult("PB", np.nan, 0.0, error="BVPS invalid")

    return results


def _confidence_relative(fd: FinancialData, multiple: str) -> float:
    conf = 0.7  # baseline: relative là sanity check, không phải primary
    if multiple == "PE" and fd.eps_ttm > 0 and len(fd.net_income) >= 3:
        # EPS ổn định?
        ni = fd.net_income[-3:]
        if min(ni) > 0 and np.std(ni)/np.mean(ni) < 0.3:
            conf *= 1.2
    if multiple == "PB" and fd.sector in ("Banking", "Insurance", "Securities"):
        conf *= 1.25  # P/B phù hợp financial
    return float(np.clip(conf, 0.1, 1.0))


# ============================================================
# TRIANGULATION (Confidence-Weighted)
# ============================================================

def triangulate(fd: FinancialData, methods: dict[str, MethodResult]) -> tuple[float, dict[str, float]]:
    """
    Confidence-weighted average của các method success.
    Weight final = base_weight × confidence, sau đó normalize về sum=1.
    """
    base_w = CFG["method_weights_base"].get(fd.sector, CFG["method_weights_base"]["default"])

    # Tính raw weight = base × confidence cho mỗi method success
    raw_weights = {}
    for m_name, m_result in methods.items():
        if not m_result.success:
            continue
        base = base_w.get(m_name, 0.0)
        if base <= 0:
            continue
        raw_weights[m_name] = base * m_result.confidence

    if not raw_weights:
        return np.nan, {}

    # Normalize
    total = sum(raw_weights.values())
    final_weights = {k: v/total for k, v in raw_weights.items()}

    # Weighted fair value
    fair_value = sum(methods[k].fair_value * w for k, w in final_weights.items())
    return fair_value, final_weights


# ============================================================
# SENSITIVITY ANALYSIS
# ============================================================

def sensitivity_grid(fd: FinancialData, wacc_base: float) -> pd.DataFrame:
    """Grid sensitivity: g × WACC → fair value FCFF (cho mã có FCFF)."""
    if fd.sector in ("Banking", "Insurance", "Securities"):
        return pd.DataFrame()
    if not fd.fcf or len(fd.fcf) < 3:
        return pd.DataFrame()

    rows = []
    for g in CFG["sens_g"]:
        row = {"g": f"{g:.1%}"}
        for dw in CFG["sens_wacc"]:
            wacc = wacc_base + dw
            CFG_backup = CFG["g_terminal"]
            CFG["g_terminal"] = g
            try:
                r = valuate_fcff(fd, wacc)
                row[f"WACC{dw:+.1%}"] = r.fair_value if r.success else np.nan
            finally:
                CFG["g_terminal"] = CFG_backup
        rows.append(row)
    return pd.DataFrame(rows)


# ============================================================
# QUALITY GATES
# ============================================================

def check_quality(fd: FinancialData) -> list[str]:
    flags = []
    g = CFG["quality_gates"]

    # ROE 3Y
    if fd.net_income and fd.total_equity:
        n = min(3, len(fd.net_income), len(fd.total_equity))
        roes = [fd.net_income[-i] / fd.total_equity[-i]
                for i in range(1, n+1) if fd.total_equity[-i] > 0]
        if roes and np.mean(roes) < g["min_roe_3y"]:
            flags.append(f"⚠️ ROE 3Y trung bình {np.mean(roes):.1%} < {g['min_roe_3y']:.0%}")

    # Debt/EBITDA
    if fd.total_debt and fd.ebit and fd.da:
        ebitda = fd.ebit[-1] + fd.da[-1] if fd.ebit[-1] is not None else 0
        if ebitda > 0:
            ratio = fd.total_debt[-1] / ebitda
            if ratio > g["max_debt_ebitda"]:
                flags.append(f"⚠️ Debt/EBITDA = {ratio:.1f}x > {g['max_debt_ebitda']}x")

    # Interest coverage
    if fd.ebit and fd.interest_expense and fd.interest_expense[-1] > 0:
        icr = fd.ebit[-1] / fd.interest_expense[-1]
        if icr < g["min_interest_cover"]:
            flags.append(f"⚠️ Interest coverage = {icr:.1f}x < {g['min_interest_cover']}x")

    # FCF âm liên tục
    if fd.fcf and len(fd.fcf) >= 3:
        neg_count = sum(1 for x in fd.fcf[-3:] if x < 0)
        if neg_count > g["max_neg_fcf_years"]:
            flags.append(f"⚠️ FCF âm {neg_count}/3 năm gần nhất")

    return flags


# ============================================================
# MAIN: VALUATE 1 TICKER
# ============================================================

def valuate_ticker(ticker: str) -> ValuationResult:
    """Định giá 1 mã. Fail-loud — mọi lỗi đều có reason."""
    log.info(f"=== [{ticker}] bắt đầu định giá ===")

    # 1. Load data
    try:
        fd = load_financial_data(ticker)
    except Exception as e:
        return ValuationResult(ticker, "default", np.nan,
                               status="FAILED", fail_reason=f"data load: {e}")

    # 2. Validate minimum data
    ok, reason = fd.has_minimum_data()
    if not ok:
        log.warning(f"[{ticker}] minimum data fail: {reason}")
        return ValuationResult(ticker, fd.sector, fd.market_price,
                               status="FAILED", fail_reason=reason)

    # 3. Quality flags
    fd.quality_flags = check_quality(fd)

    # 4. Cost of equity + WACC
    re, beta = compute_cost_of_equity(fd)
    wacc = compute_wacc(fd, re)
    log.info(f"[{ticker}] β={beta:.2f}, r_e={re:.2%}, WACC={wacc:.2%}")

    # 5. Run 3 methods
    methods = {}
    methods["FCFF"] = valuate_fcff(fd, wacc)
    methods["RI"]   = valuate_ri(fd, re)
    methods.update(valuate_relative(fd))

    # 6. Log từng method
    for m_name, m in methods.items():
        if m.success:
            log.info(f"[{ticker}] {m_name}: fair = {m.fair_value:,.0f} VND, conf = {m.confidence:.2f}")
        else:
            log.warning(f"[{ticker}] {m_name}: FAIL — {m.error}")

    # 7. Triangulate
    fair_value, weights = triangulate(fd, methods)
    if np.isnan(fair_value):
        return ValuationResult(ticker, fd.sector, fd.market_price,
                               methods=methods, quality_flags=fd.quality_flags,
                               cost_of_equity=re, beta=beta,
                               status="FAILED",
                               fail_reason="không method nào success cho sector này")

    # 8. MOS + verdict
    mos = (fair_value - fd.market_price) / fd.market_price
    if mos > CFG["mos_attractive"]:
        verdict = "🟢 HẤP DẪN"
    elif mos < CFG["mos_expensive"]:
        verdict = "🔴 ĐỊNH GIÁ CAO"
    else:
        verdict = "🟡 GẦN FAIR VALUE"

    # 9. Sensitivity
    sens = sensitivity_grid(fd, wacc)

    status = "SUCCESS" if len([m for m in methods.values() if m.success]) >= 2 else "PARTIAL"

    return ValuationResult(
        ticker=ticker, sector=fd.sector,
        market_price=fd.market_price,
        fair_value=fair_value, mos=mos, verdict=verdict,
        methods=methods, weights_applied=weights,
        quality_flags=fd.quality_flags,
        cost_of_equity=re, beta=beta,
        sensitivity=sens if not sens.empty else None,
        status=status,
    )


# ============================================================
# BATCH: ĐỌC EXCEL TỪ QUANT_PIPELINE
# ============================================================

def load_buy_signals(excel_path: str | Path) -> tuple[list[str], pd.DataFrame]:
    """
    Đọc Excel buy_summary.xlsx từ quant_pipeline.

    Schema thực tế:
        - Sheet 'Danh Sách Mua' chứa data
        - Header ở row 3 (2 dòng title phía trên)
        - Cột ticker: 'Mã' (tiếng Việt)
        - Đã filter sẵn — không cần filter thêm

    Returns:
        (tickers: list[str], signal_df: DataFrame chứa metadata gốc)
    """
    excel_path = Path(excel_path)
    if not excel_path.exists():
        raise FileNotFoundError(f"Không tìm thấy: {excel_path}")

    # --- Try to read 'Danh Sách Mua' sheet với header=3 (format chuẩn) ---
    xl = pd.ExcelFile(excel_path)
    log.info(f"Excel sheets: {xl.sheet_names}")

    target_sheet = None
    for sh in xl.sheet_names:
        sh_lower = sh.lower().replace(" ", "")
        if "danhsach" in sh_lower or "buy" in sh_lower or "danh" in sh_lower:
            target_sheet = sh
            break

    if target_sheet is None:
        target_sheet = xl.sheet_names[-1]  # fallback: sheet cuối
        log.warning(f"Không tìm thấy sheet 'Danh Sách Mua' → dùng '{target_sheet}'")

    # --- Try multiple header rows ---
    df = None
    for header_row in [3, 2, 0, 1]:
        try:
            cand = pd.read_excel(excel_path, sheet_name=target_sheet, header=header_row)
            # Check xem có cột nào trông giống ticker không
            cols_lower = [str(c).lower() for c in cand.columns]
            if any(c in ("mã", "ma", "ticker", "symbol", "code") for c in cols_lower):
                df = cand
                log.info(f"Đọc sheet '{target_sheet}' với header=row {header_row}")
                break
        except Exception:
            continue

    if df is None:
        # Fallback: dùng row 0 và cố gắng tìm
        df = pd.read_excel(excel_path, sheet_name=target_sheet)
        log.warning(f"Không tự động detect được header — dùng default")

    log.info(f"Columns thấy được: {df.columns.tolist()}")

    # --- Tìm ticker column ---
    ticker_col = None
    for cand in ["Mã", "Ma", "ticker", "symbol", "ma_ck", "Code", "TICKER", "Symbol"]:
        if cand in df.columns:
            ticker_col = cand
            break
    if ticker_col is None:
        # Tìm cột chứa data trông giống ticker (3-4 chữ in hoa)
        for col in df.columns:
            sample = df[col].dropna().astype(str).head(5).tolist()
            if sample and all(len(s) <= 5 and s.isupper() and s.isalpha() for s in sample):
                ticker_col = col
                log.info(f"Auto-detected ticker column: '{col}' (regex match)")
                break

    if ticker_col is None:
        raise KeyError(f"Không tìm thấy cột ticker. Columns: {df.columns.tolist()}")

    # --- Clean: drop NaN, dedupe ---
    df = df.dropna(subset=[ticker_col]).copy()
    df[ticker_col] = df[ticker_col].astype(str).str.strip().str.upper()
    # Lọc bỏ những row có ticker không hợp lệ (chứa dấu cách, là số, v.v.)
    df = df[df[ticker_col].str.match(r"^[A-Z]{2,5}$")]

    # --- Filter by signal nếu có cột signal (tuỳ chọn) ---
    signal_col = None
    for cand in ["signal", "Signal", "Rating", "rating", "Tín hiệu"]:
        if cand in df.columns:
            signal_col = cand
            break

    if signal_col:
        log.info(f"Found signal column '{signal_col}' — toàn bộ {len(df)} mã đã được filter sẵn từ pipeline (Score ≥ 50 & Forecast = TĂNG)")
    else:
        log.info(f"Không có cột signal — dùng toàn bộ {len(df)} mã (đã filter sẵn từ pipeline)")

    tickers = df[ticker_col].tolist()

    # --- Normalize column names cho dễ access ---
    df = df.rename(columns={ticker_col: "ticker"})

    return tickers, df


def valuate_batch(excel_path: str | Path) -> list[ValuationResult]:
    """Định giá toàn bộ mã BUY từ Excel."""
    tickers, signal_df = load_buy_signals(excel_path)
    log.info(f"=== Bắt đầu định giá {len(tickers)} mã: {tickers} ===")

    # Index by ticker để lookup nhanh
    if "ticker" in signal_df.columns:
        sig_by_ticker = signal_df.set_index("ticker").to_dict("index")
    else:
        sig_by_ticker = {}

    results = []
    for i, tk in enumerate(tickers, 1):
        log.info(f"--- [{i}/{len(tickers)}] {tk} ---")
        try:
            r = valuate_ticker(tk)
            # Attach quant signals
            r.quant_signals = sig_by_ticker.get(tk, {})
            results.append(r)
        except Exception as e:
            log.exception(f"[{tk}] hard fail batch: {e}")
            res = ValuationResult(tk, "default", np.nan,
                                  status="FAILED", fail_reason=str(e))
            res.quant_signals = sig_by_ticker.get(tk, {})
            results.append(res)

    # Summary
    n_succ = sum(1 for r in results if r.status == "SUCCESS")
    n_part = sum(1 for r in results if r.status == "PARTIAL")
    n_fail = sum(1 for r in results if r.status == "FAILED")
    log.info(f"=== Tổng kết: {n_succ} SUCCESS | {n_part} PARTIAL | {n_fail} FAILED ===")

    return results


if __name__ == "__main__":
    import sys
    if len(sys.argv) < 2:
        print("Usage: python dinhgia.py <signals_excel.xlsx>")
        sys.exit(1)
    results = valuate_batch(sys.argv[1])
    from report_cfa import build_word_report
    out = build_word_report(results)
    print(f"Report saved: {out}")
