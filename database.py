import pandas as pd
from sqlalchemy.engine import URL, create_engine
from VARIABLES import comp_mm
from azure.identity import DefaultAzureCredential
import struct

# connection parameters
server = 'tcp:ods-sql-server-us.database.windows.net'
database = 'salesops-sql-prod-us'
conn_str = (f'DRIVER=ODBC Driver 17 for SQL Server;'
            f'SERVER={server};'
            f'DATABASE={database};'
            f'ENCRYPT=yes;'
            f'TRUSTSERVERCERTIFICATE=no;'
            f'connection timeout=30')
credential = DefaultAzureCredential(exclude_interactive_browser_credential=False)
token = credential.get_token("https://database.windows.net/.default").token.encode("UTF-16-LE")
token_struct = struct.pack(f'<I{len(token)}s', len(token), token)
connection_url = URL.create("mssql+pyodbc",
                            query={"odbc_connect": conn_str})
engine = create_engine(connection_url,
                       connect_args={"attrs_before": {1256: token_struct}})


def get_queries(conn):
    """Runs all entered queries and returns their results as a dictionary called "results"
    where the query name is the key and the query results is the value."""

    sql_files_paths = {
        "tblPayout": r"C:\Users\asorensen\OneDrive - CVRx Inc\Calculate Comp\PyWorkbooks & Queries\tblPayout.sql",
        "AM": r"C:\Users\asorensen\OneDrive - CVRx Inc\Calculate Comp\PyWorkbooks & Queries\AM_INFO.sql",
        "CSR": r"C:\Users\asorensen\OneDrive - CVRx Inc\Calculate Comp\PyWorkbooks & Queries\FCE_INFO.sql",
        "comp_AM": r"C:\Users\asorensen\OneDrive - CVRx Inc\Calculate Comp\PyWorkbooks & Queries\comp_AM.sql"
    }

    sql_files = {}
    for key, path in sql_files_paths.items():
        with open(path) as file:
            sql_files[key] = file.read()

    sql_files["comp_AM"] = sql_files["comp_AM"].replace("REPLACEME", f"'2024_{comp_mm}'")
    sql_files["tblPayout"] = sql_files["tblPayout"].replace("REPLACEME", f"'2024_{comp_mm}'")

    queries = {
        "REP": sql_files["AM"],
        "CSR": sql_files["CSR"],
        "RM": "select * from qryRoster_RM",
        "tblPayout": sql_files["tblPayout"],
        "comp_AM": sql_files["comp_AM"],
        "comp_CSR": f"select * from qry_COMP_FCE_DETAIL where CLOSE_YYYYMM = '2024_{comp_mm}'",
        "comp_RM": f"select * from qry_COMP_RM_DETAIL where CLOSE_YYYYMM = '2024_{comp_mm}'"
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
