import pandas as pd
import sqlalchemy.exc
from sqlalchemy.engine import URL, create_engine
from VARIABLES import comp_mm, payout_table, tm_comp, cs_info, tm_info, cs_comp, atm_info
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
        "TM": tm_info,
        "CS": cs_info,
        "ATM": atm_info,
        "comp_TM": tm_comp,
        "comp_CS": cs_comp,
    }

    sql_files = {}
    for key, path in sql_files_paths.items():
        with open(path) as file:
            sql_files[key] = file.read()

    sql_files["comp_TM"] = sql_files["comp_TM"].replace("REPLACEME", f"'2025_{comp_mm}'")
    sql_files["comp_CS"] = sql_files["comp_CS"].replace("REPLACEME", f"'2025_{comp_mm}'")
    sql_files["tblPayout"] = sql_files["tblPayout"].replace("REPLACEME", f"'2025_{comp_mm}'")

    queries = {
        "REP": sql_files["TM"],
        "CS": sql_files["CS"],
        "ATM": sql_files["ATM"],
        "ASD": "select * from qryRoster_RM",
        "tblPayout": sql_files["tblPayout"],
        "comp_TM": sql_files["comp_TM"],
        "comp_CS": sql_files["comp_CS"],
        "comp_ASD": f"select * from qry_COMP_ASD_DETAIL where CLOSE_YYYYMM = '2025_{comp_mm}' AND SALES_COMMISSIONABLE <> 0"
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
            'TERR_NM': row['TERR_NM'],
            'THRESHOLD': row['THRESHOLD'],
            'PLAN': row['PLAN']
        }
    return info


def get_cs_names(df):
    info = {}
    for index, row in df.iterrows():
        info[row['NAME_REP']] = {
            'FNAME_REP': row['FNAME_REP'],
            'EMAIL': row['REP_EMAIL'],
            'RM_EMAIL': row['RM_EMAIL'],
            'TERR_NM': row['TERR_NM'],
            'BASE_BONUS': row['BASE_BONUS'],
            'PLAN': row['PLAN']
        }
    return info


def get_asd_names(df):
    info = {}
    for index, row in df.iterrows():
        info[row['NAME']] = {
            'FNAME': row['FNAME'],
            'LNAME': row['LNAME'],
            'EMAIL': row['EMP_EMAIL'],
            'REGION': row['REGION']
        }
    return info


def get_atm_names(df):
    info = {}
    for index, row in df.iterrows():
        info[row['NAME_REP']] = {
            'FNAME_REP': row['FNAME_REP'],
            'EMAIL': row['REP_EMAIL'],
            'RM_EMAIL': row['RM_EMAIL'],
            'TERR_NM': row['TERR_NM']
        }
    return info


class Payees:
    """Returns info on all active TM reps, CS reps, and ADSs as variables. Also returns the most recent entries in
    tblPayout as a pd.DataFrame."""

    def __init__(self):
        conn = engine.connect()
        try:
            results = get_queries(conn)
            self.tm_info = get_rep_names(results["REP"])
            self.cs_info = get_cs_names(results["CS"])
            self.atm_info = get_atm_names(results["ATM"])
            self.asd_info = get_asd_names(results["ASD"])
            self.tblpayout = results["tblPayout"]
            self.tm_comp_detail = results["comp_TM"]
            self.cs_comp_detail = results["comp_CS"]
            self.asd_comp_detail = results["comp_ASD"]
        finally:
            conn.close()
