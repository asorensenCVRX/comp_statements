WITH reps AS (
    SELECT
        A.*,
        E.NAME,
        E.DOT
    FROM
        tblRates_AM A
        LEFT JOIN tblEmployee E ON A.EID = E.[WORK E-MAIL]
    WHERE
        E.DOT IS NULL
)
SELECT
    A.NAME AS TM_NAME,
    A.EID AS TM_EID,
    B.NAME AS AM_NAME,
    B.EID AS AM_EID
FROM
    reps A
    LEFT JOIN reps B ON A.EID = B.TM_EID
WHERE
    A.[isTM?] = 1
    -- AND B.NAME IS NOT NULL;