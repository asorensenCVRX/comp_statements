SELECT
    [YYYYMM],
    [EID],
    CASE
        WHEN A.EID = 'jbeck@cvrx.com' THEN 'Josie Beck'
        ELSE isnull(B.NAME_REP, r.NAME)
    END AS [NAME],
    CASE
        WHEN A.EID = 'jbeck@cvrx.com' THEN 'MIDWEST'
        ELSE isnull(B.REGION, r.REGION)
    END AS [REGION_NM],
    B.YYYYMM_ST_DT [GUARANTEE_ST],
    B.YYYYMM_END_DT [GUARANTEE_END],
    CASE
        WHEN B.YYYYMM_END_DT >= '2023_08' THEN 1
        ELSE 0
    END [isCurrentGuar?],
    [YYYYQQ],
    CASE
        WHEN A.ROLE = 'REP' THEN 'TM'
        WHEN A.ROLE = 'FCE' THEN 'CS'
        WHEN A.ROLE = 'RM' THEN 'ASD'
        ELSE A.ROLE
    END AS [ROLE],
    CASE
        WHEN A.EID = 'jbeck@cvrx.com' THEN 'TERMED'
        ELSE isnull(B.[STATUS], r.status)
    END AS [STATUS],
    [VALUE],
    CASE
        WHEN [CATEGORY] = 'PO_AMT' THEN CAST([VALUE] AS MONEY)
        ELSE 0
    END AS [PO_AMT],
    CASE
        WHEN [CATEGORY] = 'SALES'
        AND A.ROLE IN ('REP', 'TM') THEN CAST([VALUE] AS MONEY)
        ELSE 0
    END AS [SALES],
    [CATEGORY],
    [Notes]
FROM
    [dbo].[tblPayout] A
    LEFT JOIN qryRoster B ON A.EID = B.REP_EMAIL
    AND A.YYYYMM BETWEEN DOH_YYYYMM
    AND ISNULL(DOT_YYYYMM, '2099_12')
    LEFT JOIN qryRoster_RM R ON a.EID = r.EMP_EMAIL
    AND a.ROLE IN ('RM', 'ASD')
WHERE
    A.ROLE IN('REP', 'FCE', 'RM', 'CS', 'TM', 'ASD');