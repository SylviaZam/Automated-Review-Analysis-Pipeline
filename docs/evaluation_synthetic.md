# Synthetic development set (400 answers)

| Backend | n | Theme accuracy (primary) | Theme any-match | Theme macro-F1 | Sentiment n | Sentiment accuracy | Sentiment macro-F1 | Positives called Negative |
|---|---|---|---|---|---|---|---|---|
| v1-demo | 400 | 20.0% | 20.0% | 0.16 | 56 | 37.5% | 0.46 | 0.0% |
| rules | 400 | 92.8% | 93.5% | 0.86 | 56 | 100.0% | 1.00 | 0.0% |

## v1-demo: per-theme results

| label                    |   support |   precision |   recall |   f1 |
|:-------------------------|----------:|------------:|---------:|-----:|
| brand_values             |         7 |        0.00 |     0.00 | 0.00 |
| customer_service_returns |         2 |        0.25 |     0.50 | 0.33 |
| design_style             |         1 |        0.00 |     0.00 | 0.00 |
| efficacy_results         |        70 |        0.00 |     0.00 | 0.00 |
| health_goal              |        75 |        0.00 |     0.00 | 0.00 |
| ingredients_safety       |        30 |        0.00 |     0.00 | 0.00 |
| no_concern               |        27 |        0.00 |     0.00 | 0.00 |
| other                    |        38 |        0.10 |     0.92 | 0.18 |
| payment_checkout         |         1 |        0.00 |     0.00 | 0.00 |
| price_value              |        30 |        1.00 |     0.77 | 0.87 |
| product_info             |         1 |        0.00 |     0.00 | 0.00 |
| quality                  |        25 |        1.00 |     0.32 | 0.48 |
| repeat_loyalty           |         9 |        0.00 |     0.00 | 0.00 |
| shipping_delivery        |        38 |        1.00 |     0.29 | 0.45 |
| size_fit                 |         5 |        0.67 |     0.40 | 0.50 |
| social_proof             |        28 |        0.00 |     0.00 | 0.00 |
| stock_availability       |         4 |        0.00 |     0.00 | 0.00 |
| trust_legitimacy         |         9 |        0.00 |     0.00 | 0.00 |

## v1-demo: sentiment confusion (rows = gold)

| gold     |   Mixed |   Negative |   Neutral |   Positive |
|:---------|--------:|-----------:|----------:|-----------:|
| Mixed    |       5 |          0 |         1 |          0 |
| Negative |       0 |          1 |         4 |          0 |
| Neutral  |       0 |          0 |         4 |          0 |
| Positive |       0 |          0 |        30 |         11 |

## rules: per-theme results

| label                    |   support |   precision |   recall |   f1 |
|:-------------------------|----------:|------------:|---------:|-----:|
| brand_values             |         7 |        1.00 |     0.86 | 0.92 |
| customer_service_returns |         2 |        0.25 |     0.50 | 0.33 |
| design_style             |         1 |        1.00 |     1.00 | 1.00 |
| efficacy_results         |        70 |        1.00 |     0.83 | 0.91 |
| first_time_trial         |         0 |        0.00 |     0.00 | 0.00 |
| health_goal              |        75 |        1.00 |     1.00 | 1.00 |
| ingredients_safety       |        30 |        0.71 |     0.97 | 0.82 |
| no_concern               |        27 |        1.00 |     0.93 | 0.96 |
| other                    |        38 |        0.80 |     0.87 | 0.84 |
| payment_checkout         |         1 |        1.00 |     1.00 | 1.00 |
| price_value              |        30 |        1.00 |     1.00 | 1.00 |
| product_info             |         1 |        0.00 |     0.00 | 0.00 |
| quality                  |        25 |        1.00 |     0.92 | 0.96 |
| repeat_loyalty           |         9 |        1.00 |     1.00 | 1.00 |
| shipping_delivery        |        38 |        1.00 |     1.00 | 1.00 |
| size_fit                 |         5 |        0.83 |     1.00 | 0.91 |
| social_proof             |        28 |        0.89 |     0.86 | 0.87 |
| stock_availability       |         4 |        1.00 |     1.00 | 1.00 |
| trust_legitimacy         |         9 |        0.90 |     1.00 | 0.95 |

## rules: sentiment confusion (rows = gold)

| gold     |   Mixed |   Negative |   Neutral |   Positive |
|:---------|--------:|-----------:|----------:|-----------:|
| Mixed    |       6 |          0 |         0 |          0 |
| Negative |       0 |          5 |         0 |          0 |
| Neutral  |       0 |          0 |         4 |          0 |
| Positive |       0 |          0 |         0 |         41 |
