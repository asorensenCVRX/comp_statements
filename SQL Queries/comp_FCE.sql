SELECT
    A.*
FROM
    (
        SELECT
            *,
            rank() over (
                PARTITION by OPP_ID
                ORDER BY
                    [isTarget?] DESC
            ) AS [isDUPE?]
        FROM
            qry_COMP_FCE_DETAIL
    ) AS A
WHERE
    CLOSE_YYYYMM = REPLACEME
    AND [isDUPE?] = 1
ORDER BY
    CLOSEDATE