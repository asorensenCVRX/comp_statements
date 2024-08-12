SELECT
    a.*,
    cast(b.BASE_BONUS AS decimal(10, 2)) AS BASE_BONUS,
    b.TGT_BONUS,
    Q.QUOTA
FROM
    qryRoster A
    LEFT JOIN tblFCE_COMP B ON A.REP_EMAIL = B.FCE_EMAIL
    LEFT JOIN (
        SELECT
            A.*,
            B.QUOTA
        FROM
            [dbo].[tblFCE_COMP] A
            LEFT JOIN (
                SELECT
                    A.EMAIL,
                    A.ACTIVE_YYYYMM,
                    A.DOT_YYYYMM,
                    SUM(QUOTA) [QUOTA]
                FROM
                    (
                        SELECT
                            A.EMAIL,
                            A.ACTIVE_YYYYMM,
                            A.DOT_YYYYMM,
                            ISNULL(
                                R.Quota,
                                y.QUOTA_FY
                            ) QUOTA
                        FROM
                            qryAlign_FCE A
                            LEFT JOIN tblRates_RM R ON R.REGION_ID = A.[KEY]
                            LEFT JOIN qryRates_AM Y ON a.[KEY] = Y.TERR_ID
                        WHERE
                            [TYPE] IN('REGION', 'TERR')
                    ) AS A
                GROUP BY
                    A.EMAIL,
                    A.ACTIVE_YYYYMM,
                    A.DOT_YYYYMM
            ) AS B ON a.FCE_EMAIL = b.EMAIL
    ) AS Q ON A.REP_EMAIL = Q.FCE_EMAIL
WHERE
    a.[ROLE] = 'FCE'
    AND a.[isLATEST?] = 1
    AND (
        DOT IS NULL
        OR DOT >= DATEADD(MONTH, DATEDIFF(MONTH, 0, GETDATE()), 0)
    )
    AND TENNURE_MONTHS > 0;