import pandas as pd
import sqlalchemy.exc
from sqlalchemy.engine import URL, create_engine
from VARIABLES import comp_mm, payout_table, am_comp, fce_info, am_info, tm_reports, fce_comp
from azure.identity import DefaultAzureCredential
import struct
from pprint import pprint

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
        "tblPayout": payout_table,
        "AM": am_info,
        "CSR": fce_info,
        "comp_AM": am_comp,
        "comp_FCE": fce_comp,
        "TM_reports": tm_reports
    }

    sql_files = {}
    for key, path in sql_files_paths.items():
        with open(path) as file:
            sql_files[key] = file.read()

    sql_files["comp_AM"] = sql_files["comp_AM"].replace("REPLACEME", f"'2024_{comp_mm}'")
    sql_files["comp_FCE"] = sql_files["comp_FCE"].replace("REPLACEME", f"'2024_{comp_mm}'")
    sql_files["tblPayout"] = sql_files["tblPayout"].replace("REPLACEME", f"'2024_{comp_mm}'")

    queries = {
        "REP": sql_files["AM"],
        "CSR": sql_files["CSR"],
        "RM": "select * from qryRoster_RM",
        "tblPayout": sql_files["tblPayout"],
        "comp_AM": sql_files["comp_AM"],
        "comp_CSR": sql_files["comp_FCE"],
        "comp_RM": f"select * from qry_COMP_RM_DETAIL where CLOSE_YYYYMM = '2024_{comp_mm}'",
        "TM_reports": sql_files["TM_reports"]
    }

    results = {}
    for key, query in queries.items():
        print(f"Fetching {key} info...")
        try:
            results[key] = pd.read_sql_query(query, conn)
        except sqlalchemy.exc.ProgrammingError:
            pprint(f"Error with {key}: \n {query}")
            input("stop")
        else:
            print("Success")

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


def get_tm_reports(df):
    info = {}
    for index, row in df.iterrows():
        if row['TM_EID'] not in info:
            info[row['TM_EID']] = []
        info[row['TM_EID']].append(row['AM_EID'])
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
            self.tm_reports = get_tm_reports(results["TM_reports"])
        finally:
            conn.close()
