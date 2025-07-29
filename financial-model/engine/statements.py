from .revenue import project_revenue
from .costs import project_costs

def create_3_statements(inputs, years=5):
    revenue = project_revenue(inputs["starting_revenue"], inputs["revenue_growth"], years)
    costs = project_costs(revenue, inputs["margin"])
    profit = [r - c for r, c in zip(revenue, costs)]

    # Cash flow: assume all profit converts to cash (simplified)
    cash_flow = profit

    # Balance sheet (simplified): debt stays constant
    debt = [inputs["debt"]] * years
    cash = []
    running_cash = 0
    for cf in cash_flow:
        running_cash += cf
        cash.append(running_cash)

    assets = cash
    equity = [a - d for a, d in zip(assets, debt)]

    return {
        "income_statement": {"Year": list(range(1, years+1)), "Revenue": revenue, "Costs": costs, "Profit": profit},
        "cash_flow": {"Year": list(range(1, years+1)), "Cash Flow": cash_flow},
        "balance_sheet": {"Year": list(range(1, years+1)), "Assets": assets, "Debt": debt, "Equity": equity}
    }
