from sqlalchemy import *
import pandas as pd
import re
import time
import os
os.environ['NLS_LANG'] = 'SIMPLIFIED CHINESE_CHINA.UTF8'


def get_db(db_str):
    if db_str == 'scb5':
        db = create_engine('')
    elif db_str == 'scb6':
        db = create_engine('')
    else:
        db = None
    return db


def sql_2_excel(filename, sql, db_str):
    db = get_db(db_str)
    t0 = time.time()
    df = pd.read_sql(sql, db)
    df.to_excel(filename, index=False)
    t1 = time.time()
    print('cost ', t1 - t0)


def excel_2_sql(filename, tablename, db_str):
    t0 = time.time()
    df = pd.read_excel(filename)
    t1 = time.time()
    print("read cost", (t1 - t0))

    if 'IDX' in df.columns or 'idx' in df.columns:
        pass
    else:
        df.index.rename('idx', inplace=True)
        df.reset_index(inplace=True)
    df_2_sql(df, tablename, db_str)


def df_2_sql(df, tablename, db_str):
    t0 = time.time()
    cols = []
    dtype = {}
    for d in df.columns:
        s = re.sub(r'\W', '', str(d))
        if not s:
            s = 'blank'
        if len(s) > 10:
            s = s[:10]
        if re.search(r'^\d', s):
            s = 'a' + s
        suff = 1
        tmp = s
        while s in cols:
            s = tmp + str(suff)
            suff += 1
        cols.append(s)
        if df.dtypes[d] in ('object', 'uint64'):
            df[d].fillna('', inplace=True)
            df[d] = df[d].astype(str)
            dtype[s] = String(1000)
    df.columns = cols
    db = get_db(db_str)

    try:
        db.execute('drop table ' + tablename + ' purge')
    except Exception as e:
        print(str(e))
    try:
        df.to_sql(tablename, db, index=False, dtype=dtype, chunksize=1000)
        print("success:", tablename)
    except Exception as e:
        print(str(e))
    t1 = time.time()
    print("import cost", (t1 - t0))
