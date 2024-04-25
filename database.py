import pandas as pd
from sqlalchemy.engine import URL, create_engine

# connection parameters
server = 'CAP-MAXtX8vWsQc\\SQLEXPRESS'
database = 'CVRx'
conn_str = f'DRIVER=ODBC Driver 17 for SQL Server;SERVER={server};DATABASE={database};Trusted_Connection=yes'
connection_url = URL.create("mssql+pyodbc", query={"odbc_connect": conn_str})
engine = create_engine(connection_url)


def get_rep_names(role: str) -> dict:
    """role must equal either 'FCE' or 'REP'"""
    # connect to database
    conn = engine.connect()
    # execute query and save to df
    query = f"select * from qryRoster where [ROLE] = '{role}' and [STATUS] = 'ACTIVE' and [isLATEST?] = 1"
    df = pd.read_sql_query(query, conn)
    # Close the connection
    conn.close()
    info = {}
    for index, row in df.iterrows():
        info[row['NAME_REP']] = {
            'EMAIL': row['REP_EMAIL'],
            'RM_EMAIL': row['RM_EMAIL'],
            'TERR_NM': row['TERR_NM']
        }
    # info = {NAME: {EMAIL: value, RM_EMAIL: value, TERR_NM: value}, ...}
    return info


def get_rm_names() -> dict:
    """Returns a dictionary of RM names and info."""
    # connect to database
    conn = engine.connect()
    # execute query and save to df
    query = "select * from qryRoster_RM"
    result = pd.read_sql_query(query, conn)
    # rm_df = pd.DataFrame(result)
    # Close the connection
    conn.close()
    info = {}
    for index, row in result.iterrows():
        info[row['NAME']] = {
            'EMAIL': row['EMP_EMAIL'],
            'REGION': row['REGION']
        }
    # info = {NAME: {EMAIL: value, RM_EMAIL: value, TERR_NM: value}, ...}
    return info


def get_tblpayout() -> pd.DataFrame:
    """Returns all the most recent entries from tblPayout as a pd.dataframe."""
    # connect to database
    conn = engine.connect()
    # execute query and save to df
    query = "select * from tblPayout where YYYYMM = (select max(tblPayout.YYYYMM) from tblPayout)"
    result = pd.read_sql_query(query, conn)
    conn.close()
    return result


def get_comp_detail(role: str) -> pd.DataFrame:
    """Returns all the most recent entries from the comp detail tables as a pd.dataframe."""
    # connect to database
    conn = engine.connect()
    # execute query and save to df
    if role.lower() == 'rep':
        query = ('select * from qry_COMP_AM_DETAIL where CLOSE_YYYYMM = (select max(CLOSE_YYYYMM) from '
                 'qry_COMP_AM_DETAIL)')
        result = pd.read_sql_query(query, conn)
        conn.close()
        return result
    elif role.lower() == 'fce':
        query = ('select * from qry_COMP_FCE_DETAIL where CLOSE_YYYYMM = (select max(CLOSE_YYYYMM) from '
                 'qry_COMP_FCE_DETAIL)')
        result = pd.read_sql_query(query, conn)
        conn.close()
        return result
    elif role.lower() == 'rm':
        query = ('select * from qry_COMP_RM_DETAIL where CLOSE_YYYYMM = (select max(CLOSE_YYYYMM) from '
                 'qry_COMP_RM_DETAIL)')
        result = pd.read_sql_query(query, conn)
        conn.close()
        return result


class Payees:
    """Returns info on all active reps, fces, and rms as variables. Also returns the most recent entries in tblPayout
    as a pd.DataFrame."""

    def __init__(self):
        self.am_info: dict = get_rep_names('REP')
        self.fce_info: dict = get_rep_names('FCE')
        self.rm_info: dict = get_rm_names()
        self.tblpayout: pd.DataFrame = get_tblpayout()
        self.am_comp_detail: pd.DataFrame = get_comp_detail('REP')
        self.fce_comp_detail: pd.DataFrame = get_comp_detail('FCE')
        self.rm_comp_detail: pd.DataFrame = get_comp_detail('RM')
