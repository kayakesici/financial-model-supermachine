def project_interest(debt, interest_rate, years=5):
    return [debt * interest_rate for _ in range(years)]
