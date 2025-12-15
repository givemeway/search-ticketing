import sqlite3
import requests
import os
import pickle
from bs4 import BeautifulSoup
headers = {
    'User-Agent': "Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:138.0) Gecko/20100101 Firefox/138.0"
}
login_data = {
    'do': "",
    'userid': "",
    'passwd': "",
    'submit': ''
}

request_errors = ['error-http', 'error-connection',
                  'error-timeout', 'error-request']


def create_connection(db):
    conn = None
    try:
        if os.path.exists("DB"):
            conn = sqlite3.connect(f'DB/{db}')
        else:
            os.makedirs("DB")
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
     DATE TEXT,
     ATTACHMENT TEXT
     );''')
    conn.execute('PRAGMA journal_mode = WAL;')


def create_emailNotFound_table(conn):
    conn.execute('''CREATE TABLE IF NOT EXISTS emailnotfound
    ( ID INTEGER PRIMARY KEY NOT NULL,
      EMAIL TEXT,
      SUBJECT TEXT,
      DATE TEXT,
      ATTACHMENT TEXT
    );''')
    conn.execute('PRAGMA journal_mode = WAL;')


def create_subjectNotFound_table(conn):
    conn.execute('''CREATE TABLE IF NOT EXISTS subjectnotfound
    ( ID INTEGER PRIMARY KEY NOT NULL,
      EMAIL TEXT,
      SUBJECT TEXT,
      DATE TEXT,
      ATTACHMENT TEXT,
      MATCH INT
    );''')
    conn.execute('PRAGMA journal_mode = WAL;')


def verify2FA(sessionData, _session):
    global headers
    s = sessionData['session']
    del sessionData['session']
    url = "https://ticket.idrive.com/scp/login.php"
    try:
        r = s.post(url, data=sessionData, headers=headers)
        loginPageSoup = BeautifulSoup(r.content, 'html5lib')
        if r.status_code == 403:
            return (403, "Login Forbidden")  # unauthorized
        status, msg = validateLogin(r, s)

        if status == 200:
            _session.emit(s)
            print("saving the session to pickle")
            with open('session.pickle', 'wb') as f:
                pickle.dump(s.cookies, f)
            return (200, "SUCCESS", "")
        elif status == 401 and isinstance(msg, dict):
            label_2fa = loginPageSoup.select(
                "div.field-label.required")[0].find("label").get("for")
            CSRF = loginPageSoup.find(
                'input', attrs={'name': '__CSRFToken__'})['value']
            login_data = {}
            login_data['__CSRFToken__'] = CSRF
            login_data[label_2fa] = ""
            login_data['sessionID'] = label_2fa
            login_data['do'] = "2fa"
            login_data['session'] = s
            return (401, "Invalid Code", login_data)
        elif status == 401 and isinstance(msg, str):
            query = '''DELETE FROM login WHERE rowid=1'''
            conn = create_connection('ticket.db')
            with conn:
                cur = conn.cursor()
                cur.execute(query)
            return (401, msg, "")
    except requests.exceptions.HTTPError as errh:
        return (request_errors[0], errh)
    except requests.exceptions.ConnectionError as errc:
        return (request_errors[1], errc)
    except requests.exceptions.Timeout as errt:
        return (request_errors[2], errt)
    except requests.exceptions.RequestException as err:
        return (request_errors[3], err)


def ticket_session(name, secret, _session, _2fa, CSRFToken, isExpired=False):
    global headers
    with requests.Session() as s:
        url = "https://ticket.idrive.com/scp/login.php"
        try:
            r = s.get(url, headers=headers)
            if r.status_code == 403:
                return (403, "Login Forbidden")  # unauthorized
            else:
                try:
                    with open('session.pickle', 'rb') as f:
                        s.cookies.update(pickle.load(f))
                        status, msg = validateLogin(r, s)
                        if status == 200:
                            _session.emit(s)
                            return (200, "SUCCESS")
                        else:
                            os.remove('session.pickle')
                            return build_header(r, s, url, name, secret, _session, _2fa, CSRFToken)

                except Exception as e:
                    print("Session file is empty, creating a new session.")
                    return build_header(r, s, url, name, secret, _session, _2fa, CSRFToken)
        except requests.exceptions.HTTPError as errh:
            return (request_errors[0], errh)
        except requests.exceptions.ConnectionError as errc:
            return (request_errors[1], errc)
        except requests.exceptions.Timeout as errt:
            return (request_errors[2], errt)
        except requests.exceptions.RequestException as err:
            return (request_errors[3], err)

# __CSRFToken__: 8e5f3894030dd6faf627e8ec5769bb3a0dbb32e6
# b97458938d61be: 406603
# do: 2fa
# ajax: 1


# <input type="hidden" name="__CSRFToken__" value="5f6a94748bd0482250277f5454ac7f393a2cab34" /><div class="form-simple">
#             <div class="flush-left custom-field" id="field_d1693da4947952" >
#         <div>
#           <div class="field-label required">
#         <label for="d1693da4947952">
#             Verification Code:
#               <span class="error">*</span>
#           </label>
#         </div>

#         __CSRFToken__: 5f6a94748bd0482250277f5454ac7f393a2cab34
# d1693da4947952: 512044
# do: 2fa
# ajax: 1

def validateLogin(r, session, username=None, password=None):
    loginPageSoup = BeautifulSoup(r.content, 'html5lib')
    loginMsg = loginPageSoup.find(id="login-message")
    print("Login Message is -- ", loginMsg)
    if loginMsg is not None and len(loginMsg) > 0:
        loginMsg = loginMsg.getText().strip().lower()
        if loginMsg == "Invalid login".lower() or loginMsg == "Authentication Required".lower():
            query = '''DELETE FROM login WHERE rowid=1'''
            conn = create_connection('ticket.db')
            with conn:
                cur = conn.cursor()
                cur.execute(query)
            return (401, loginMsg)

        elif loginMsg == "Access denied".lower():
            query = '''DELETE FROM login WHERE rowid=1'''
            conn = create_connection('ticket.db')
            with conn:
                cur = conn.cursor()
                cur.execute(query)
            return (401, loginMsg)
        elif loginMsg == "2FA Pending".lower():
            label_2fa = loginPageSoup.select(
                "div.field-label.required")[0].find("label").get("for")
            CSRF = loginPageSoup.find(
                'input', attrs={'name': '__CSRFToken__'})['value']
            login_data = {}
            login_data['__CSRFToken__'] = CSRF
            login_data[label_2fa] = ""
            login_data['sessionID'] = label_2fa
            login_data['do'] = "2fa"
            login_data['session'] = session
            updateDB(username, password)
            return (401, login_data)
        elif loginMsg == "Invalid Code".lower():
            return (401, loginMsg)
    else:
        return (200, 'Success')


def updateDB(username, password):
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
                                        VALUES(?,?)''', (username, password))
                cur.connection.commit()
            cur.execute("SELECT * FROM mbox")
            rows = cur.fetchall()
            if rows is not None:
                query = '''DELETE FROM mbox WHERE rowid=?'''
                for row in rows:
                    if row[5] is None:
                        val = (row[0],)
                        cur.execute(query, val)
                cur.connection.commit()


def build_header(r, s, url, username, password, _session, _2fa, CSRFToken):
    global login_data
    try:
        soup = BeautifulSoup(r.content, 'html5lib')
        login_data['__CSRFToken__'] = soup.find(
            'input', attrs={'name': '__CSRFToken__'})['value']
        login_data['userid'] = username
        login_data['passwd'] = password
        CSRFToken.emit(login_data['__CSRFToken__'])
        r = s.post(url, data=login_data, headers=headers)
        status, data = validateLogin(r, s, username, password)
        print("Status is ", status, " and data is ", data)

        if status == 200:
            updateDB(username, password)
            with open('session.pickle', 'wb') as f:
                print("saving the session to pickle")
                pickle.dump(s.cookies, f)
            _session.emit(s)
            return (200, "SUCCESS")
        elif status == 401 and isinstance(data, dict):
            _2fa.emit(data)
            return (401, "2FA")
        elif status == 401 and isinstance(data, str):
            query = '''DELETE FROM login WHERE rowid=1'''
            conn = create_connection('ticket.db')
            with conn:
                cur = conn.cursor()
                cur.execute(query)
            return (401, data)

    except Exception as e:
        # er.exception(f'DB Exception-2: {e}')
        print(e)
        return ('DB Exception-2', e)


if __name__ == "__main__":
    pass
