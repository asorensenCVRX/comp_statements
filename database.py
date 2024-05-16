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
    queries = {
        "REP": "select * from qryRoster where [ROLE] = 'REP' and [STATUS] = 'ACTIVE' and [isLATEST?] = 1",
        "FCE": "select * from qryRoster where [ROLE] = 'FCE' and [STATUS] = 'ACTIVE' and [isLATEST?] = 1",
        "RM": "select * from qryRoster_RM",
        "tblPayout": "select * from tblPayout where YYYYMM = (select max(tblPayout.YYYYMM) from tblPayout)",
        "comp_AM": "select * from qry_COMP_AM_DETAIL where CLOSE_YYYYMM = (select max(CLOSE_YYYYMM) from "
                   "qry_COMP_AM_DETAIL) and [isSale?] = 1",
        "comp_FCE": "select * from qry_COMP_FCE_DETAIL where CLOSE_YYYYMM = (select max(CLOSE_YYYYMM) from "
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
            'EMAIL': row['REP_EMAIL'],
            'RM_EMAIL': row['RM_EMAIL'],
            'TERR_NM': row['TERR_NM']
        }
    return info


def get_rm_names(df):
    info = {}
    for index, row in df.iterrows():
        info[row['NAME']] = {
            'EMAIL': row['EMP_EMAIL'],
            'REGION': row['REGION']
        }
    return info


class Payees:
    """Returns info on all active reps, fces, and rms as variables. Also returns the most recent entries in tblPayout
    as a pd.DataFrame."""

    def __init__(self):
        conn = engine.connect()
        try:
            results = get_queries(conn)
            self.am_info = get_rep_names(results["REP"])
            self.fce_info = get_rep_names(results["FCE"])
            self.rm_info = get_rm_names(results["RM"])
            self.tblpayout = results["tblPayout"]
            self.am_comp_detail = results["comp_AM"]
            self.fce_comp_detail = results["comp_FCE"]
            self.rm_comp_detail = results["comp_RM"]
        finally:
            conn.close()
