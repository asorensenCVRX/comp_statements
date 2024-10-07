WITH rolecounts AS (
    SELECT
        YYYYMM,
        EID,
        count(DISTINCT ROLE) AS ROLE_COUNT
    FROM
        TBLpAYOUT
    GROUP BY
        YYYYMM,
        EID
)
SELECT
    T.*,
    RC.ROLE_COUNT
FROM
    tblPayout T
    JOIN rolecounts RC ON T.YYYYMM = RC.YYYYMM
    AND T.EID = RC.EID
WHERE
    RC.ROLE_COUNT > 1
ORDER BY
    EID,
    YYYYMM;