def dcf_valuation(cash_flows, discount_rate, exit_multiple):
    npv = 0
    for i, cf in enumerate(cash_flows):
        npv += cf / ((1 + discount_rate) ** (i + 1))

    terminal_value = cash_flows[-1] * exit_multiple
    terminal_pv = terminal_value / ((1 + discount_rate) ** len(cash_flows))

    enterprise_value = npv + terminal_pv
    return enterprise_value
