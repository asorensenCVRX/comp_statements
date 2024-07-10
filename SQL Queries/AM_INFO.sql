SELECT
    *
FROM
    qryRoster
WHERE
    [ROLE] = 'REP'
    AND [isLATEST?] = 1
    AND (
        DOT IS NULL
        OR DOT >= DATEADD(MONTH, DATEDIFF(MONTH, 0, GETDATE()), 0)
    )
    AND TENNURE_MONTHS > 0