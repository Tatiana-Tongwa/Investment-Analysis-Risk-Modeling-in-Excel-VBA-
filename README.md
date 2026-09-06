# Investment-Analysis-Risk-Modeling-in-Excel-VBA
A reusable capital-investment analysis model for evaluating a proposed soft-drink production plant. The project combines free-cash-flow modeling, capital budgeting, scenario analysis, sensitivity analysis, break-even analysis, and Monte Carlo simulation in a manager-friendly Excel tool.

Portfolio focus: financial modeling, investment analysis, Excel/VBA automation, risk analysis, and management reporting.



Business case

The model evaluates a greenfield production plant with a two-year construction period and a 12-year operating life. The base case uses the following assumptions:

Assumption

Base case

Initial investment

$10.0m

Construction period

2 years

Nominal discount rate

10%

Inflation

4%

Corporate tax rate

20%

Initial annual volume

1.0m bottles

Volume after three operating years

1.5m bottles

Sales price

$2.00 / bottle

Production cost

$0.50 / bottle

Operating life

12 years

Depreciation

Straight-line over 13 years

Working-capital investment

$50k annually for the first 5 operating years, released over the last 5 years

Terminal value

Scrap value based on remaining depreciated value

What the model does

Builds a full operating and investment cash-flow schedule.

Separates accounting profit from actual cash flow using EBITDA, EBIT, taxes, depreciation, capital expenditure, working capital, and terminal value.

Calculates NPV, IRR, Profitability Index, Payback Period, and Discounted Payback Period.

Automates discounted payback so it updates when assumptions change.

Provides dashboard charts for cash flow, discounted cash flow, cumulative cash flow, EBITDA/EBIT, and salvage value.

Runs break-even analysis by solving for input values that drive NPV to zero.

Runs sensitivity analysis for two profitability outputs across five major assumptions.

Includes scenario analysis for alternative combinations of tax, production cost, and inflation.

Includes a Monte Carlo simulation with five uncertain inputs and uniform distributions.

Produces a printable management report and includes VBA-driven refresh/export controls.

Base-case results

Metric

Result

Net Present Value

$4.74m

Internal Rate of Return

12.27%

Profitability Index

1.47x

Payback Period

8.17 years

Discounted Payback Period

9.87 years

The base case creates value: NPV is positive, IRR is above the 10% nominal hurdle rate, and PI is above 1. The main caution is the relatively long payback period, which reflects the two-year construction phase and back-loaded operating cash flows.

Break-even analysis

The model calculates the threshold for each selected assumption at which NPV becomes zero.

Variable

Base case

NPV break-even

Approx. headroom

Initial investment

$10.00m

$15.84m

+58%

Sales price / bottle

$2.00

$1.45

-28%

Volume after three years

1.50m

0.78m

-48%

Inflation

4.0%

-2.02%

-6.0 pp

Nominal discount rate

10.0%

16.76%

+6.8 pp

These thresholds turn the model from a valuation worksheet into a decision tool: managers can see how much adverse movement the project can absorb before value creation disappears.

Sensitivity analysis

The model tests NPV and Profitability Index against relative changes from -50% to +50% in:

Initial investment

Sales price per bottle

Nominal discount rate

Inflation rate

Sales volume after three years

Within the tested ranges, investment cost, long-run sales volume, and the discount rate are especially influential on NPV. The analysis also shows which levers management can actively influence (construction cost, pricing, volume) versus market/financial variables that are mainly monitored (inflation and discount rate).

Scenario analysis

Three scenario bundles are included around tax, unit production cost, and inflation. One useful modeling insight is that scenario labels such as "optimistic" and "pessimistic" should be interpreted as bundles of assumptions, not automatically as higher/lower NPV states. Because the model is nominal, inflation materially affects the nominal revenue and cost path, so combinations can produce non-obvious rankings.

Monte Carlo simulation

The simulation treats five variables as uncertain:

Initial investment

Sales price per bottle

Sales volume after three years

Inflation rate

Nominal discount rate

For each simulation, the workbook draws values from user-defined uniform ranges and records investment outcomes. The tool is designed to summarize expected NPV, minimum/maximum NPV, dispersion, downside probability, and an NPV distribution.

For portfolio presentation, run the simulation with a substantially larger number of trials than a quick test run (for example, 1,000+), then refresh the report before taking screenshots.

VBA automation

The workbook contains controls linked to VBA routines for:

break-even refresh (Breakeven_analysis_macro)

sensitivity refresh (Sensitivity)

Monte Carlo simulation (macro_monte_carlo)

report/PDF export (print_report)

See vba/README.md for the source-code publishing plan.

Model architecture

The workbook is organized into separate model components:

Core model / inputs / calculations

Sensitivity analysis

Scenario summary

Monte Carlo assumptions and output data

Monte Carlo histogram

Management report

A detailed walkthrough is available in docs/methodology.md.

Repository note

The original course brief and the original macro-enabled workbook are intentionally not included in this public-ready package yet. The source workbook contains personal document metadata and a hard-coded local export path. Those should be sanitized before the .xlsm file is committed publicly. The dashboard image in this repository is a rendered, metadata-free preview.

Skills demonstrated

Financial modeling: capital budgeting, free cash flow, NPV, IRR, PI, payback, discounted payback, terminal value, working capital
Risk analysis: break-even, sensitivity, scenarios, Monte Carlo simulation
Excel: formulas, charts, management dashboard, report design
Automation: VBA macros, refresh controls, PDF export workflow
