SELECT
    YYYYMM,
    EID,
    YYYYQQ,
    ROLE,
    [STATUS],
    cast(value AS decimal(15, 2)) AS 'VALUE',
    CATEGORY,
    Notes
FROM
    tblPayout
WHERE
    YYYYMM = REPLACEME;


/***** TROUBLESHOOTING *****/
-- SELECT
--     YYYYMM,
--     EID,
--     YYYYQQ,
--     ROLE,
--     [STATUS],
--     value,
--     CATEGORY,
--     Notes
-- FROM
--     tblPayout
-- WHERE
--     TRY_CAST(value AS decimal(15, 2)) IS NULL
--     AND value IS NOT NULL
--     AND YYYYMM = '2024_06'