SELECT
    R.*,
    Q.THRESHOLD,
    Q.[PLAN],
    RD.RECOVERABLE_DRAW
FROM
    qryRoster R
    LEFT JOIN (
        SELECT
            YYYYQQ,
            TERRITORY_ID,
            EID,
            SUM(THRESHOLD) AS THRESHOLD,
            SUM([PLAN]) AS [PLAN]
        FROM
            qryQuota_Monthly
        WHERE
            /* get last month's quarter */
            YYYYQQ = (
                FORMAT(DATEADD(MONTH, -1, GETDATE()), 'yyyy') + '_Q' + CAST(
                    CEILING(MONTH(DATEADD(MONTH, -1, GETDATE())) / 3.0) AS VARCHAR
                )
            )
        GROUP BY
            YYYYQQ,
            TERRITORY_ID,
            EID
    ) AS Q ON R.REP_EMAIL = Q.EID
    LEFT JOIN (
        SELECT
            EMP_EMAIL,
            SUM(RECOVERABLE_DRAW) AS RECOVERABLE_DRAW
        FROM
            qryRecoverable_Draw
        GROUP BY
            EMP_EMAIL
    ) RD ON RD.EMP_EMAIL = R.REP_EMAIL
WHERE
    [ROLE] = 'REP' -- AND [isLATEST?] = 1
    AND ISNULL(DOT_YYYYMM, '2099_12') >= FORMAT(DATEADD(MONTH, -1, GETDATE()), 'yyyy_MM')
    AND DOH_YYYYMM < FORMAT(GETDATE(), 'yyyy_MM')