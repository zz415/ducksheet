-- Creates two views from the loaded sales + products tables

CREATE OR REPLACE VIEW sales_summary AS
SELECT
    region,
    category,
    CAST(EXTRACT(YEAR  FROM order_date) AS INTEGER) AS order_year,
    CAST(EXTRACT(MONTH FROM order_date) AS INTEGER) AS order_month,
    COUNT(*)                                                      AS order_count,
    SUM(quantity)                                                 AS total_qty,
    ROUND(SUM(revenue), 2)                                        AS total_revenue,
    ROUND(AVG(revenue), 2)                                        AS avg_order_value,
    COUNT(*) FILTER (WHERE is_renewal = true)                     AS renewal_count
FROM sales
WHERE status = 'Closed Won'
GROUP BY region, category, order_year, order_month;


CREATE OR REPLACE VIEW rep_leaderboard AS
SELECT
    rep_name,
    COUNT(*)                                                                      AS total_orders,
    COUNT(*) FILTER (WHERE status = 'Closed Won')                                AS won,
    ROUND(SUM(revenue) FILTER (WHERE status = 'Closed Won'), 2)                  AS won_revenue,
    ROUND(AVG(revenue) FILTER (WHERE status = 'Closed Won'), 2)                  AS avg_deal_size,
    ROUND(
        COUNT(*) FILTER (WHERE status = 'Closed Won') * 100.0 / COUNT(*), 1
    )                                                                             AS win_rate_pct
FROM sales
GROUP BY rep_name
ORDER BY won_revenue DESC;
