def project_revenue(start_value, growth_rate, years=5):
    revenue = []
    current = start_value
    for _ in range(years):
        revenue.append(current)
        current *= (1 + growth_rate)
    return revenue
