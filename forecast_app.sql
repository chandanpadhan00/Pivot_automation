BEGIN;

WITH src AS (
    SELECT product, '2025-01'::text AS month, jan_25 AS val FROM asnfdm.forecasted_application_2025
    UNION ALL SELECT product, '2025-02', feb_25 FROM asnfdm.forecasted_application_2025
    UNION ALL SELECT product, '2025-03', mar_25 FROM asnfdm.forecasted_application_2025
    UNION ALL SELECT product, '2025-04', apr_25 FROM asnfdm.forecasted_application_2025
    UNION ALL SELECT product, '2025-05', may_25 FROM asnfdm.forecasted_application_2025
    UNION ALL SELECT product, '2025-06', jun_25 FROM asnfdm.forecasted_application_2025
    UNION ALL SELECT product, '2025-07', jul_25 FROM asnfdm.forecasted_application_2025
    UNION ALL SELECT product, '2025-08', aug_25 FROM asnfdm.forecasted_application_2025
    UNION ALL SELECT product, '2025-09', sep_25 FROM asnfdm.forecasted_application_2025
    UNION ALL SELECT product, '2025-10', oct_25 FROM asnfdm.forecasted_application_2025
    UNION ALL SELECT product, '2025-11', nov_25 FROM asnfdm.forecasted_application_2025
    UNION ALL SELECT product, '2025-12', dec_25 FROM asnfdm.forecasted_application_2025
)

UPDATE asnfdm.forecast_table ft
SET forecasted_application = src.val
FROM src
WHERE ft.vendor = 'REGALORX'
  AND ft.drug = src.product
  AND ft.month = src.month
  AND ft.month LIKE '2025-%';

COMMIT;