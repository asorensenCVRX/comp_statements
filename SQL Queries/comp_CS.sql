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
    )
ORDER BY
    CLOSEDATE