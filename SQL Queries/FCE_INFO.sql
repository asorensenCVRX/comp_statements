SELECT
    a.*,
    cast(b.BASE_BONUS AS decimal(10, 2)) AS BASE_BONUS,
    b.TGT_BONUS
FROM
    qryRoster A
    LEFT JOIN tblFCE_COMP B ON A.REP_EMAIL = B.FCE_EMAIL
WHERE
    a.[ROLE] = 'FCE'
    AND a.[isLATEST?] = 1
    AND (
        DOT IS NULL
        OR DOT >= DATEADD(MONTH, DATEDIFF(MONTH, 0, GETDATE()), 0)
    )
    AND TENNURE_MONTHS > 0