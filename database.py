import pandas as pd
from sqlalchemy.engine import URL, create_engine

# connection parameters
server = 'ods-sql-server-us.database.windows.net'
database = 'salesops-sql-prod-us'
username = 'asorensen@cvrx.com'
conn_str = (f'DRIVER=ODBC Driver 17 for SQL Server;SERVER={server};DATABASE={database};UID={username};'
            f'Authentication=ActiveDirectoryInteractive')
connection_url = URL.create("mssql+pyodbc", query={"odbc_connect": conn_str})
engine = create_engine(connection_url)


def get_queries(conn):
    sql_files = {}
    with open(r"C:\Users\asorensen\OneDrive - CVRx Inc\Calculate Comp\PyWorkbooks & Queries\tblPayout.sql") as file:
        sql_files["tblPayout"] = file.read()
    with open(r"C:\Users\asorensen\OneDrive - CVRx Inc\Calculate Comp\PyWorkbooks & Queries\FCE_INFO.sql") as file:
        sql_files["CSR"] = file.read()
    queries = {
        "REP": "select * from qryRoster where [ROLE] = 'REP' and [STATUS] = 'ACTIVE' and [isLATEST?] = 1",
        "CSR": sql_files["CSR"],
        "RM": "select * from qryRoster_RM",
        "tblPayout": sql_files["tblPayout"],
        "comp_AM": "select * from qry_COMP_AM_DETAIL where CLOSE_YYYYMM = (select max(CLOSE_YYYYMM) from "
                   "qry_COMP_AM_DETAIL) and [isSale?] = 1",
        "comp_CSR": "select * from qry_COMP_FCE_DETAIL where CLOSE_YYYYMM = (select max(CLOSE_YYYYMM) from "
                    "qry_COMP_FCE_DETAIL)",
        "comp_RM": "select * from qry_COMP_RM_DETAIL where CLOSE_YYYYMM = (select max(CLOSE_YYYYMM) from "
                   "qry_COMP_RM_DETAIL)"
    }

    results = {}
    for key, query in queries.items():
        print(f"Fetching {key} info...")
        results[key] = pd.read_sql_query(query, conn)

    return results


def get_rep_names(df):
    info = {}
    for index, row in df.iterrows():
        info[row['NAME_REP']] = {
            'FNAME_REP': row['FNAME_REP'],
            'EMAIL': row['REP_EMAIL'],
            'RM_EMAIL': row['RM_EMAIL'],
            'TERR_NM': row['TERR_NM']
        }
    return info


def get_csr_names(df):
    info = {}
    for index, row in df.iterrows():
        info[row['NAME_REP']] = {
            'FNAME_REP': row['FNAME_REP'],
            'EMAIL': row['REP_EMAIL'],
            'RM_EMAIL': row['RM_EMAIL'],
            'TERR_NM': row['TERR_NM'],
            'BASE_BONUS': row['BASE_BONUS']
        }
    return info


def get_rm_names(df):
    info = {}
    for index, row in df.iterrows():
        info[row['NAME']] = {
            'FNAME': row['FNAME'],
            'EMAIL': row['EMP_EMAIL'],
            'REGION': row['REGION']
        }
    return info


class Payees:
    """Returns info on all active reps, csrs, and rms as variables. Also returns the most recent entries in tblPayout
    as a pd.DataFrame."""

    def __init__(self):
        conn = engine.connect()
        try:
            results = get_queries(conn)
            self.am_info = get_rep_names(results["REP"])
            self.csr_info = get_csr_names(results["CSR"])
            self.rm_info = get_rm_names(results["RM"])
            self.tblpayout = results["tblPayout"]
            self.am_comp_detail = results["comp_AM"]
            self.csr_comp_detail = results["comp_CSR"]
            self.rm_comp_detail = results["comp_RM"]
        finally:
            conn.close()
