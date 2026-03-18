-- Revenue breakdown by region + category, ranked by total revenue
-- Runs against the sales_summary view built by build_views.sql

SELECT
    region,
    category,
    SUM(order_count)                                AS total_orders,
    SUM(total_qty)                                  AS total_units,
    ROUND(SUM(total_revenue), 2)                    AS total_revenue,
    ROUND(AVG(avg_order_value), 2)                  AS avg_order_value,
    SUM(renewal_count)                              AS total_renewals,
    ROUND(
        SUM(renewal_count) * 100.0 / SUM(order_count), 1
    )                                               AS renewal_rate_pct
FROM sales_summary
GROUP BY region, category
ORDER BY total_revenue DESC;
