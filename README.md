# 📊 Apple Inc. (AAPL) — Financial Model & DCF Valuation

**Amogh Gupta | M.Sc. Quantitative Finance, Kiel University, Germany**

---

I built this project to bridge the gap between theory and practice. My Quantitative Finance programme covers valuation in depth, but I wanted to actually sit down and build a full 3-statement model and DCF from scratch — the kind of work you'd do in an IB internship. Apple felt like the right company to start with: clean financials, a well-documented business, and an interesting valuation story given the Services segment transformation.

The result is an 8-sheet Excel model with 5 years of historical data and 5-year projections, a DCF valuation, trading comps against 5 peers, and a sensitivity analysis across 42 WACC/growth rate combinations. I also rebuilt the entire model in Python using openpyxl, partly to automate it and partly because working through it programmatically forced me to understand every single formula.

My base case implies a share price of ~$172 against a reference price of $195 — roughly 12% downside — which puts Apple in **Hold** territory at current levels.

---

## 🏢 Apple — Why It's Interesting Right Now

Apple is not a simple hardware story anymore. Services (~22% of revenue, ~70%+ gross margins) is growing at ~15% annually and is steadily shifting the margin profile of the whole company upward. That's the bull case in one sentence.

The bear case is China (~19% of revenue), App Store regulatory risk, and the question of whether Apple Intelligence actually drives a meaningful iPhone upgrade cycle or whether it's just marketing.

| Detail | Value |
|--------|-------|
| **Ticker** | AAPL (NASDAQ) |
| **Sector** | Technology |
| **FY End** | September 30 |
| **Reference Price** | $195.00 |
| **Market Cap** | ~$3.0 Trillion |
| **Shares Outstanding** | ~15.4 Billion |

---

## 📐 Key Assumptions

The most important judgment calls in any model are the revenue growth rates and the WACC. Here's my thinking:

### Revenue Growth

| Year | Growth | Reasoning |
|------|--------|-----------|
| FY2024E | 4.3% | iPhone 15 cycle recovery after FY2023 dip |
| FY2025E | 7.0% | iPhone 16 + early AI feature adoption |
| FY2026E | 8.0% | Services compounding + India market opening up |
| FY2027E | 7.5% | Vision Pro ecosystem, wearables maturing |
| FY2028E | 6.5% | Normalisation — harder to grow off a larger base |

### Margins

| Metric | FY2023A | FY2028E | Driver |
|--------|---------|---------|--------|
| Gross Margin | 44.1% | 45.5% | Services mix shift |
| EBITDA Margin | 32.8% | 32.8% | Stable opex leverage |
| Net Margin | 25.3% | 27.0% | Interest income on net cash position |

### WACC

| Component | Value |
|-----------|-------|
| Risk-Free Rate | 4.25% (10Y UST) |
| Equity Risk Premium | 5.50% (Damodaran) |
| Beta | 1.25 (5Y monthly vs S&P 500) |
| Cost of Equity | 11.1% |
| After-Tax Cost of Debt | 4.06% |
| Equity Weight | 92% |
| **WACC** | **10.0%** |

---

## 📊 DCF — Methodology & Output

I used Unlevered Free Cash Flow (UFCF) discounted at WACC, which is standard for a company like Apple where the capital structure is relatively stable.

```
EBIT − Taxes on EBIT + D&A − CapEx − ΔNWC = UFCF
```

For terminal value I used two methods and averaged them:
- **Gordon Growth Model** at 3.0% terminal growth rate
- **Exit Multiple** at 22.0x EV/EBITDA (peer median)

### Output

| Component | Value |
|-----------|-------|
| PV of FCFs (FY24E–28E) | $450B |
| PV of Terminal Value | $2,280B |
| **Enterprise Value** | **$2,730B** |
| Less: Net Debt | ($66B) |
| **Equity Value** | **$2,664B** |
| **Implied Share Price** | **~$172** |

Terminal value is ~83% of EV — high, but expected for a business with Apple's long-term FCF profile.

### Sensitivity — Implied Share Price vs WACC & Terminal Growth

|  | WACC 8.5% | WACC 9.5% | **WACC 10.0%** | WACC 10.5% | WACC 11.5% |
|--|-----------|-----------|----------------|------------|------------|
| **TGR 2.0%** | $215 | $181 | $167 | $155 | $135 |
| **TGR 3.0%** | $248 | $204 | **$172** | $158 | $136 |
| **TGR 4.0%** | $306 | $241 | $213 | $190 | $157 |

---

## 🔍 Trading Comps

| Peer | EV/Revenue | EV/EBITDA |
|------|-----------|-----------|
| Microsoft | 12.1x | 24.3x |
| Alphabet | 6.2x | 18.4x |
| Meta | 7.8x | 22.1x |
| Samsung | 2.1x | 14.2x |
| Sony | 1.8x | 12.3x |
| **Median** | **6.5x** | **21.6x** |
| Apple implied | ~$159/share | ~$169/share |

---

## 💡 Conclusion: HOLD

At $195, Apple is not obviously cheap. My blended valuation range is $149–$194 with a midpoint of ~$171 — implying ~12% downside at current prices.

That said, I think the market is pricing in a reasonable bull case on Services and AI. If Apple Intelligence genuinely accelerates the upgrade cycle and Services reaches 30%+ of revenue by FY2027, the bull case of ~$230 is achievable.

The position I'd take: **Hold existing, don't add at $195.** Wait for a pullback toward $160–$170 before building a position.

> ⚠️ Student project — not investment advice.

---

## 📚 Data Sources

| Data | Source |
|------|--------|
| Historical financials | Apple 10-K filings, SEC EDGAR |
| Beta / price data | Yahoo Finance |
| Equity risk premium | Damodaran Online |
| Peer multiples | Public filings |

---

*M.Sc. Quantitative Finance — Kiel University | Built as part of IB internship preparation*
