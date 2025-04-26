SELECT
    *
FROM
    qry_COMP_CS_DETAIL
WHERE
    CLOSE_YYYYMM = REPLACEME
    AND (
        (
            SALES_COMMISSIONABLE <> 0
            AND STAGENAME = 'Revenue Recognized'
        )
        OR (
            TGT_PO_YYYYMM = FORMAT(DATEADD(MONTH, -1, GETDATE()), 'yyyy_MM')
        )
    )
ORDER BY
    CLOSEDATE