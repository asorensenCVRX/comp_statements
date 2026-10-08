SELECT
    *
FROM
    qryRoster_RM
WHERE
    FORMAT(START_DT, 'yyyy_MM') < FORMAT(DATEADD(MONTH, -1, GETDATE()), 'yyyy_MM')
    AND EMP_EMAIL <> 'open@cvrx.com'