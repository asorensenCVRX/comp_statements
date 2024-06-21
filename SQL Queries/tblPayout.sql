SELECT
    YYYYMM,
    EID,
    YYYYQQ,
    ROLE,
    [STATUS],
    cast(value AS decimal(10, 2)) AS 'VALUE',
    CATEGORY,
    Notes
FROM
    tblPayout
WHERE
    YYYYMM = REPLACEME