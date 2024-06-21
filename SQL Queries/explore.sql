SELECT
    AM_EMAIL,
    [IsCasePaid?],
    OPP_OWNER_NAME,
    'REP' AS ROLE,
    REGION_NM,
    0 AS ISIMPL,
    CPAS_PA_SUB_DT,
    CONCAT('CASE# ', CASENUMBER),
    OPP_ID,
    NULL,
    FACILITY_NAME,
    ACT_ID,
    NULL,
    NULL,
    REP_PO_YYYYMM,
    NULL,
    NULL,
    'CPAS',
    NULL,
    NULL,
    NULL,
    NULL,
    NULL,
    NULL,
    NULL,
    NULL,
    NULL,
    NULL,
    NULL,
    NULL,
    NULL,
    NULL,
    PHYSICIAN,
    NULL,
    NULL,
    NULL
FROM
    qry_COMP_CPAS_SPIF;


SELECT
    *
FROM
    qry_COMP_CPAS_SPIF
WHERE
    AM_EMAIL = 'jclemmons@cvrx.com'