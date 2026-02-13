WITH src AS (
    SELECT product, '2025-01'::text AS month, jan_25 AS val FROM asnfdm.forecasted_application_2025
    UNION ALL SELECT product, '2025-02', feb_25 FROM asnfdm.forecasted_application_2025
)
SELECT ft.drug, ft.month, ft.published_quarter,
       ft.forecasted_application AS old_value,
       src.val AS new_value
FROM asnfdm.forecast_table ft
JOIN src
  ON ft.drug = src.product
 AND ft.month = src.month
WHERE ft.vendor = 'REGALORX'
ORDER BY ft.drug, ft.month, ft.published_quarter;