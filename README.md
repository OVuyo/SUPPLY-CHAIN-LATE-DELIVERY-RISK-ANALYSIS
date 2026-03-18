# Supply Chain Late Delivery Risk Prediction & Shipping Cost Optimization

**An end-to-end analytics project built entirely in Microsoft Excel — from exploratory analysis through machine learning, Monte Carlo simulation, and constrained optimization.**

Uses real supply chain data (180k+ orders from DataCo Global) to predict late deliveries, quantify cost exposure, and find the optimal shipping mode allocation.

---

## The Problem

DataCo Global has a late delivery rate of **54.8%** — more than half of all orders arrive late. This drives penalty costs, customer churn, and operational inefficiency. The company uses four shipping modes (Standard Class, First Class, Second Class, Same Day), but the current allocation is not cost-optimized.

The question: **Can we predict which orders will be late before they ship, quantify the monthly cost risk, and find a better shipping mode mix?**

## The Approach

| Phase | Method | Sheet |
|-------|--------|-------|
| Exploration | Summary statistics, pivot breakdowns, conditional formatting | `1_Data_EDA` |
| Feature Engineering | Binary encoding, derived features, median-based buckets | `2_Feature_Engineering` |
| Prediction | Logistic regression (gradient descent, 3,000 iterations) | `3_Logistic_Regression` |
| Risk Quantification | Monte Carlo simulation (1,000 scenarios, NORMINV + RAND) | `4_Monte_Carlo_Sim` |
| Optimization | Excel Solver (GRG Nonlinear, constrained allocation) | `5_Solver_Optimization` |
| Communication | Executive dashboard with KPIs and recommendations | `6_Dashboard` |

## Key Results

- **70.8% prediction accuracy** using only pre-shipment features (no data leakage)
- **First Class shipping** has a coefficient of +3.89 in the logistic model — selecting it nearly guarantees a late delivery
- **Same Day shipping** has a coefficient of -2.70 — it actively reduces late delivery risk
- Monthly cost exposure averages **~$99k** with a **95th percentile VaR of ~$128k**
- Solver optimization can reduce the weighted late rate from **54.5% to under 45%**, saving an estimated **$5–15k/month**

## Data

**Source:** [DataCo SMART SUPPLY CHAIN FOR BIG DATA ANALYSIS](https://data.mendeley.com/datasets/8gx2fvg2k6/3) (Mendeley Data, CC BY 4.0)

- 180,519 orders from DataCo Global
- 53 original features covering provisioning, production, sales, and distribution
- Products: Clothing, Sports, Electronics
- Markets: Europe, LATAM, Pacific Asia, USCA, Africa
- Sampled to 8,000 orders (stratified) for Excel workability

**Citation:**
Constante, F., Silva, F., & Pereira, A. (2019). DataCo SMART SUPPLY CHAIN FOR BIG DATA ANALYSIS. Mendeley Data, V3. doi: 10.17632/8gx2fvg2k6.3

## On Data Leakage

This matters. The dataset includes `Days for shipping (real)` and `Delivery Status` — both are **post-hoc** variables that are only known after an order has already been delivered. Using them would give you near-perfect accuracy (the day difference deterministically encodes lateness), but that model is useless in production because you can't know actual shipping days before the order ships.

I deliberately excluded these features. The model uses only what's available at order time: scheduled shipping days, product price, quantity, sales amount, discount rate, shipping mode, customer segment, and market. This drops accuracy from ~100% to 70.8%, but it's a model you could actually deploy.

## Workbook Structure

```
Supply_Chain_Optimization_Completed.xlsx
│
├── 1_Data_EDA
│   ├── 8,000 cleaned order records (15 columns)
│   ├── Summary statistics panel
│   ├── Late rate by shipping mode / segment / market
│   └── Bar chart visualization
│
├── 2_Feature_Engineering
│   ├── 18 model-ready features (all formula-linked to Sheet 1)
│   ├── Binary encoding: 4 shipping modes → 3 dummies
│   ├── Binary encoding: 3 segments → 2 dummies
│   ├── Binary encoding: 5 markets → 4 dummies
│   └── Derived: price bucket, high quantity flag
│
├── 3_Logistic_Regression
│   ├── 18 trained coefficients (gradient descent)
│   ├── Sigmoid function: P(Late) = 1/(1+exp(-z))
│   ├── Confusion matrix: TP=2392, FP=326, TN=3274, FN=2008
│   ├── Metrics: Accuracy=70.8%, Precision=88.0%, Recall=54.4%
│   └── 100-row prediction sample with live formulas
│
├── 4_Monte_Carlo_Sim
│   ├── 10 editable assumptions (blue cells)
│   ├── 1,000 simulated monthly scenarios
│   ├── Random demand via NORMINV(RAND(), mean, stdev)
│   ├── Late rate draws with variability
│   └── Risk summary: mean, VaR 95%, VaR 99%, P(cost > $100k)
│
├── 5_Solver_Optimization
│   ├── 4 shipping modes with cost and late rate parameters
│   ├── Decision variables (yellow cells): allocation percentages
│   ├── Current vs Optimized cost comparison
│   ├── Constraints: sum=100%, min 5%, max 80%, late rate ≤ 45%
│   └── Step-by-step Solver setup instructions
│
└── 6_Dashboard
    ├── 4 KPIs: Late Rate, Accuracy, VaR 95%, Savings
    ├── 5 key findings
    ├── Actionable recommendation
    └── Cost breakdown table (linked to Sheet 5)
```

## How to Use

1. **Open in Excel** (not Google Sheets — Solver and NORMINV won't work there)
2. **Press F9** to regenerate Monte Carlo random draws
3. **Go to Sheet 5** → Data tab → Solver
4. Set up Solver per the on-sheet instructions
5. Click Solve → the yellow cells update → savings populate across Sheets 5 and 6

## Technical Notes

- **152,566 live formulas**, zero hardcoded calculation values
- All calculations are dynamic — change an assumption, everything recalculates
- Logistic regression coefficients were trained via gradient descent in Python, then placed as blue input cells in Excel. The prediction formulas (sigmoid, threshold) are native Excel
- Monte Carlo uses Excel's built-in NORMINV and RAND functions — no VBA, no add-ins
- Solver uses GRG Nonlinear method for the constrained optimization

## Tools

- Microsoft Excel (Solver add-in)
- Python (data sampling and coefficient training only — not required to use the workbook)

## What I'd Do Next

- Add a sensitivity analysis sheet — vary the late penalty cost and see how the optimal allocation shifts
- Build a scenario comparison table (best case / base case / worst case)
- Connect to live order data via Power Query for real-time risk monitoring
- Test additional classifiers (decision tree, k-NN) and compare in a model selection sheet

---

*Built as a portfolio project demonstrating the full analytics lifecycle: exploration → prediction → risk quantification → optimization → communication.*
