SELECT
    a.external_source,
    a.case_status,
    a.enrollment_status,
    COUNT(DISTINCT NULLIF(BTRIM(a.patient_id), '')) AS distinct_patients
FROM asnfdm.f_pcd2_detailed_report_vw a
WHERE a.case_created_date >= DATE '2023-10-01'
  AND UPPER(BTRIM(a.case_status)) NOT IN ('CANCELLED', 'NOT APPROVED')
  AND NULLIF(BTRIM(a.patient_id), '') IS NOT NULL
  AND NOT EXISTS (
      SELECT 1
      FROM asnfdm.d_phi dp
      WHERE NULLIF(BTRIM(dp.patient_id), '') IS NOT NULL
        AND BTRIM(dp.patient_id) = BTRIM(a.patient_id)
  )
GROUP BY
    a.external_source,
    a.case_status,
    a.enrollment_status
ORDER BY
    a.external_source, a.case_status, a.enrollment_status;