import sqlite3,requests
from bs4 import BeautifulSoup
headers = {
        'User-Agent': "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/96.0.4664.45 Safari/537.36"
    }
login_data = {
        'do': "scplogin",
        'userid': "sandeep.kumar@idrive.com",
        'passwd': "",
        'submit': ''
        }

request_errors = [ 'error-http','error-connection','error-timeout','error-request']

def create_connection(db):
    conn = None
    try:
        conn = sqlite3.connect(f'DB/{db}')
    except Exception as e:
        pass
        # er.exception("DB exception-1: {e}")
    return conn

def create_login_table(conn):
    conn.execute('''CREATE TABLE IF NOT EXISTS login
    (USERNAME TEXT NOT NULL,
     PASSWORD TEXT NOT NULL
    );''')
    conn.execute('PRAGMA journal_mode = WAL;')

def create_error_table(conn):
    conn.execute('''CREATE TABLE IF NOT EXISTS error
    (ID INTEGER PRIMARY KEY NOT NULL,
    EMAIL TEXT,
    SUBJECT TEXT,
    DATE TEXT);''')
    conn.execute('PRAGMA journal_mode = WAL;')

def create_mbox_table(conn):
    conn.execute('''CREATE TABLE IF NOT EXISTS mbox
    (ID INTEGER PRIMARY KEY NOT NULL,
     DATE TEXT,  
     ITEMS INTEGER,
     START INTEGER,
     END INTEGER,
     EMAILS_START INTEGER,
     EMAILS_END INTEGER,
     SUBJECT_START INTEGER,
     SUBJECT_END INTEGER
     );''')
    conn.execute('PRAGMA journal_mode = WAL;')

def create_email_table(conn):
    conn.execute('''CREATE TABLE IF NOT EXISTS emails
    (ID INTEGER PRIMARY KEY NOT NULL,
     EMAIL TEXT,
     SUBJECT TEXT,
     DATE TEXT
     );''')
    conn.execute('PRAGMA journal_mode = WAL;')

def create_emailNotFound_table(conn):
    conn.execute('''CREATE TABLE IF NOT EXISTS emailnotfound
    ( ID INTEGER PRIMARY KEY NOT NULL,
      EMAIL TEXT,
      SUBJECT TEXT,
      DATE TEXT
    );''')
    conn.execute('PRAGMA journal_mode = WAL;')

def create_subjectNotFound_table(conn):
    conn.execute('''CREATE TABLE IF NOT EXISTS subjectnotfound
    ( ID INTEGER PRIMARY KEY NOT NULL,
      EMAIL TEXT,
      SUBJECT TEXT,
      DATE TEXT
    );''')
    conn.execute('PRAGMA journal_mode = WAL;')


def ticket_session(name,secret,_session):
    global headers
    with requests.Session() as s:
        url = "https://ticket.idrive.com/scp/login.php"
        try:
            r = s.get(url,headers=headers)
            if r.status_code == 403:
                return 403 # unauthorized
            else:
                return build_header(r,s,url,name,secret,_session)  
        except requests.exceptions.HTTPError as errh:
                return (request_errors[0],errh)
        except requests.exceptions.ConnectionError as errc:
                return (request_errors[1],errc)
        except requests.exceptions.Timeout as errt:
                return (request_errors[2],errt)
        except requests.exceptions.RequestException as err:
                return (request_errors[3],err)

def build_header(r,s,url,username,password,_session):
                global login_data
                try:
                    soup = BeautifulSoup(r.content,'html5lib')
                    login_data['__CSRFToken__'] = soup.find('input',attrs={'name':'__CSRFToken__'})['value']
                    login_data['userid']= username
                    login_data['passwd']= password
                    r = s.post(url,data=login_data,headers=headers)
                    if r.status_code == 200:
                        conn = create_connection('ticket.db')
                        create_login_table(conn)
                        create_email_table(conn)
                        create_emailNotFound_table(conn)
                        create_subjectNotFound_table(conn)
                        create_error_table(conn)
                        create_mbox_table(conn)
                        with conn: 
                            if conn is not None:
                                cur = conn.cursor()
                                cur.execute("SELECT * FROM login")
                                row = cur.fetchone()
                                if row is None:
                                    cur.execute('''INSERT INTO login(USERNAME,PASSWORD) 
                                                VALUES(?,?)''',(username,password))
                                    cur.connection.commit()
                                cur.execute("SELECT * FROM mbox")
                                rows = cur.fetchall()
                                if rows is not None:
                                    query ='''DELETE FROM mbox WHERE rowid=?'''
                                    for row in rows:
                                        if row[5] is None:
                                            val = (row[0],)
                                            cur.execute(query,val)
                                    cur.connection.commit()
                        _session.emit(s)
                        return 200 #login success
                    else:
                        
                        query ='''DELETE FROM login WHERE rowid=1'''
                        conn = create_connection('ticket.db')
                        with conn:
                            cur = conn.cursor()
                            cur.execute(query)
                        return 422 # login failed

                except Exception as e:
                    # er.exception(f'DB Exception-2: {e}')
                    return ('DB Exception-2',e)

if __name__=="__main__":
    pass

