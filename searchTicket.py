from login import create_connection
from concurrent.futures import ThreadPoolExecutor
from threading import Event
from bs4 import BeautifulSoup
import concurrent.futures
from parse import error,missing_sub_logger,missing_ticket,logger
import requests,re,os,sys
from login import headers
import urllib.parse
import math
from sklearn.feature_extraction.text import TfidfVectorizer
from sklearn.metrics.pairwise import cosine_similarity

idx=0
email_idx = 0
subject_idx = 0
err_idx = 0
n_workers = 5
ticket_reply_reg = re.compile(r'''\[#[I|i][D|d][0-9]{9}\]|\[#[I|i][B|b][0-9]{9}\]|\[#[R|r][P|p][0-9]{9}\]''',re.MULTILINE)
ticket_num_reg =re.compile(r'''[I|i][D|d][0-9]{9}|[I|i][B|b][0-9]{9}|[R|r][P|p][0-9]{9}''',re.MULTILINE)
subject_regex = re.compile(r''''[#ID[0-9]+]''',re.MULTILINE)
idriveinc_domain_regex = re.compile(r'''([a-zA-Z0-9.!#$%&'*+-/=?^_`{|}~-]+@[a-zA-Z0-9.!#$%&'*+-/=?^_`{|}~-]+(?:\.idrive\.com|\.remotepc\.com|\.ibackup\.com|\.idrivecompute\.com|\.idrivemirror\.com))''',re.MULTILINE)

searchURL_sorted ='https://ticket.idrive.com/scp/tickets.php?sort=date&dir=0&a=search&search-type=&query='

exclude_list = []

def get_path(filename):
    if hasattr(sys, "_MEIPASS"):
        return os.path.join(sys._MEIPASS, filename)
    else:
        return filename  

def fetchExcludeList():
    exclude_list = []
    try:
        if os.path.exists('logs/exclude.log'):
            with open('logs/exclude.log','r') as f:
                lines = f.readlines()
                for line in lines:
                    exclude_list.append(line.split('\n')[0])
        else:

            log_file = get_path('logs/exclude.log')
            with open(log_file,'r') as f:
                lines = f.readlines()
                for line in lines:
                    exclude_list.append(line.split('\n')[0])
                with open('logs/exclude.log','w') as f:
                    for line in lines:     
                        f.write(line)
        return exclude_list
    except Exception as e:
        error.exception(f'File Error: Exclude file not found. Using Default - {e}')
        return None

def ticketing(_dict):
    global n_workers
    global subject_idx,email_idx,idx,err_idx
    global exclude_list

    idx=0
    email_idx = 0
    subject_idx = 0
    err_idx = 0
    exclude_list = fetchExcludeList()
    event = Event()
    abnormal = False
    with ThreadPoolExecutor(n_workers) as exe:
        futures = [exe.submit(fetch_ticket,[mail],_dict,event) for mail in _dict["emails"]]
        for future in concurrent.futures.as_completed(futures):
            if event.is_set():
                abnormal = True
                future.cancel()
            else:
                future.result()
    query_mbox = '''SELECT * FROM mbox ORDER BY ID DESC LIMIT 1'''
    query_emails = '''SELECT * FROM emailnotfound ORDER BY ID DESC LIMIT 1'''
    query_subject = '''SELECT * FROM subjectnotfound ORDER BY ID DESC LIMIT 1'''
    conn = create_connection('ticket.db')
    if conn is not None:
        with conn:
            cur = conn.cursor()
            cur.execute(query_mbox)
            mbox_cur = cur.fetchone()
            cur.execute(query_emails)
            email_cur = cur.fetchone()
            cur.execute(query_subject)
            subject_cur = cur.fetchone()
            if email_cur is not None:
                val = (int(email_cur[0])-email_idx,int(email_cur[0]),int(mbox_cur[0]))
                cur.execute('''UPDATE mbox SET EMAILS_START=?, EMAILS_END=? WHERE ID=?''',val)
            if subject_cur is not None:
                val = (int(subject_cur[0])-subject_idx,int(subject_cur[0]),int(mbox_cur[0]))
                cur.execute('''UPDATE mbox SET SUBJECT_START=?, SUBJECT_END=? WHERE ID=?''',val)
            cur.connection.commit()
    _dict['warning'].emit(abnormal)

    # print("Info","Search Complete!")

def fetch_ticket(chunk_tickets,_dict,event):
    global idx,subject_idx,email_idx,exclude_list,err_idx
    if event.is_set():
        idx = idx+1
        _dict['processing'].emit((chunk_tickets[0][1],idx))
        return
    # idx=0
    # email_idx = 0
    # subject_idx = 0

    email_query ='''INSERT INTO emailnotfound(EMAIL,SUBJECT,DATE,ATTACHMENT)
                    VALUES(?,?,?,?) '''
    sub_query ='''INSERT INTO subjectnotfound(EMAIL,SUBJECT,DATE,ATTACHMENT,MATCH)
            VALUES(?,?,?,?,?)'''
    conn = create_connection('ticket.db')
    # temp_error = None
    with conn:
        cur = conn.cursor()
        for ticket in chunk_tickets:
            idx = idx+1
            _dict['processing'].emit((ticket[1],idx))
            if ticket is not None:
                for item in ticket[1]:
                    if len(item)>0:

                        if item.lower() not in exclude_list:
                            
                            r = search_query(_dict['session'],searchQuery=item)
                            if isinstance(r,str):
                                if str(r).startswith('error-'):
                                    error.exception(f"{r}: {ticket}")
                                    err_idx+=1
                                    temp = ticket
                                    temp[0] = err_idx 
                                    _dict['error'].emit((f'{r}',temp, item))
                                    event.set()
                                    # temp_error = True
                                    # break
                            elif isinstance(r,Exception):
                                err_idx+=1
                                temp = ticket
                                temp[0] = err_idx
                                _dict['error'].emit((f'{r}',temp, item))
                                error.exception(f'Unknown Error: {r}') 
                                # temp_error = True
                                # break
                                
                            elif r.status_code == 200:
                                try:
                                    decoded_content = r.content.decode('utf-8')
                                    soup = BeautifulSoup(r.content,'html5lib')
                                    search = decoded_content.__contains__('There are no tickets matching your criteria.')
                                    #build the page with the search results
                                    status = extract_search(soup)
                                    # no results for email
                                    if search:
                                        if ticket[2] is not None:
                                            val = "Level 1 : Email: "+ item + " Subject: " + ticket[2]
                                        else:
                                            val = "Level 1 : Email: "+ item + " Subject: " + " "
                                        if len(re.findall(idriveinc_domain_regex,val))==0:
                                            missing_ticket.debug(val)
                                            value = (item,ticket[2],ticket[3],f'''{ticket[4]}''')
                                            cur.execute(email_query,value)
                                            cur.connection.commit()
                                            temp = ticket
                                            email_idx +=1
                                            temp[0] = email_idx
                                            _dict['emailNotFound'].emit(temp)

                                    else:
                                    # if results found match the subject lines
                                    # check if subject is empty
                                        try:
                                            if ticket[2] is not None:
                                                sub_obj = compare_tickets(status,ticket[2])
                                            else:
                                                sub_obj = compare_tickets(status," ")
                                            if sub_obj==True:
                                                if ticket[2] is not None:
                                                    val = "Level 2: Email: "+ item + " Subject: " + ticket[2]
                                                else:
                                                    val = "Level 2: Email: "+ item + " Subject: " + " "
                                                
                                                if len(re.findall(idriveinc_domain_regex,val))==0:
                                                        matchProbability = extractCustomerMessages(_dict['session'],status,ticket[5])
                                                    # if not extractCustomerMessages(_dict['session'],status,ticket[5]):
                                                        missing_ticket.debug(val)
                                                        value = (item,ticket[2],ticket[3],f'''{ticket[4]}''',matchProbability)
                                                        cur.execute(sub_query,value)
                                                        cur.connection.commit()
                                                        temp = ticket
                                                        subject_idx +=1
                                                        temp[0] = subject_idx
                                                        temp.insert(2,matchProbability)
                                                        _dict['subjectNotFound'].emit(temp)

                                            elif isinstance(sub_obj,list):
                                                for tick in sub_obj:
                                                    r = search_query(_dict['session'],searchQuery=tick)
                                                    ticket_soup = BeautifulSoup(r.content,'html5lib')
                                                    ticket_found = extract_search(ticket_soup)
                                                    if not len(ticket_found):
                                                        if ticket[2] is not None:
                                                            val = "Level 3: Email: "+ item + " Subject: " + ticket[2]
                                                        else:
                                                            val = "Level 3: Email: "+ item + " Subject: " + " "
                                                        if len(re.findall(idriveinc_domain_regex,val))==0:
                                                            missing_ticket.debug(val) 
                                                            matchProbability = extractCustomerMessages(_dict['session'],status,ticket[5])
                                                            value = (item,ticket[2],ticket[3],f'''{ticket[4]}''',matchProbability)
                                                            cur.execute(sub_query,value)
                                                            cur.connection.commit()
                                                            temp = ticket
                                                            subject_idx +=1
                                                            temp[0] = subject_idx
                                                            temp.insert(2,matchProbability)
                                                            _dict['subjectNotFound'].emit(temp)
                                        except Exception as e:
                                            error.exception(f"Error-Level 2: {e} : {ticket}") 

                                except Exception as e:
                                    error.exception("Error-Level 1: {} : {}".format(e,ticket))

                            elif str(r.status_code).startswith('4'):
                                # print.showerror("Error", "Unauthorized access 403")
                                error.exception(f"Error-Unauthorized: {ticket}")
                                err_idx+=1
                                temp = ticket
                                temp[0] = err_idx
                                _dict['error'].emit(('403-Unauthorized',temp, item))
                                event.set()
                                # temp_error = True
                                # break
                            elif str(r.status_code).startswith('5'):
                                # print("Error", "Ticket Error 500")
                                error.exception(f"Error-500: {ticket}")
                                err_idx+=1
                                temp = ticket
                                temp[0] = err_idx
                                _dict['error'].emit(('500-Ticket Error',temp, item))
                                # temp_error = True
                                # break
                            elif str(r.status_code).startswith('3'):
                                # resource not found - logged out
                                error.exception(f"Error-InvalidSession: {ticket}")
                                err_idx+=1
                                temp = ticket
                                temp[0] = err_idx
                                _dict['error'].emit(('Invalid Session',temp, item))
                                event.set()
                                # temp_error = True
                                # break

            per = round((idx/int(_dict['emailCount']))*100)   
            _dict["pBar"].emit(per)
            # if temp_error:
            #     break

def search_query(s,ticketURL=None,searchQuery=None):
    # query = '\"{}\"'.format(urllib.parse.quote_plus(searchQuery.strip()))
    # url = searchURL_sorted + query
    try:
        if searchQuery is not None:
            query = '\"{}\"'.format(urllib.parse.quote_plus(searchQuery.strip()))
            url = searchURL_sorted + query
        else:
            url = ticketURL
        r  = s.get(url,headers=headers)
        return r
        # r  = s.get(url,headers=headers)
        # return r
    except requests.exceptions.HTTPError as errh:
        error.exception(f'Request-HTTPError: {errh} {searchQuery}')
        return 'error-http'
    except requests.exceptions.ConnectionError as errc:
        error.exception(f'Request-ConnectionError: {errc} {searchQuery}')
        return 'error-connection'
    except requests.exceptions.Timeout as errt:
        error.exception(f'Request-Timeout: {errt} {searchQuery}')
        return 'error-timeout'
    except requests.exceptions.RequestException as err:
        error.exception(f'Request-Exception: {err} {searchQuery}')
        return 'error-requestionException'
    except Exception as e:
        error.exception(f'Search_Error: {e}')     
        return e

def findSimilarityScore(doc1,doc2):

    vectorizer = TfidfVectorizer()

    # Compute the TF-IDF vectors for the documents
    tfidf_matrix = vectorizer.fit_transform([doc1, doc2])

    # Compute the cosine similarity between the documents
    similarity = cosine_similarity(tfidf_matrix[0:1], tfidf_matrix)
    
    return similarity

def extractCustomerMessages(session,tickets,email_data):
    minProb = float("-inf")
    for content in email_data.read_email_payload():
        # if content[0] == "text/plain":
        text2 = content[2]
        for ticket in tickets.values():
            try:
                r  = session.get(ticket['ticketUrl'],headers=headers)
                soup = BeautifulSoup(r.content,'html5lib')
                for _,response in enumerate(soup.find_all("div",class_="thread-entry message avatar")):
                    for body in response.find_all("div",class_="thread-body"):
                        text1 = body.text.strip()
                        similarity = findSimilarityScore(text1,text2)
                        minProb = max(similarity[0][1],minProb)
                        # if float(similarity[0][1]) >= 0.6:
                        #     return True 
            except Exception as e:
                error.exception(e)

    return 0.0 if math.isinf(minProb) else minProb*100
    
def extract_search(soup):
    list_obj ={}
    ticketHomeurl = "https://ticket.idrive.com/scp/"
    for count,tr in enumerate(soup.find_all('tr')):
        search_obj={}
        for i,td in enumerate(tr.find_all('td')):
            if i==1:
                search_obj['ticket'] = td.text.strip()
                for a in td.find_all('a'):
                    search_obj['ticketUrl'] = ticketHomeurl + a.get('href')
            elif i==2:
                search_obj['created']= td.text.strip()
            elif i==3:
                children = td.findChildren('div',recursive=True)
                for child in children:
                    search_obj['subject'] = child.text.strip()
            else:
                continue
        if len(search_obj):       
            list_obj[count] = search_obj
        
    return list_obj
def compare_tickets(results,email_subject):
    for item in results.values():
        parsed_email_subject = split_subject(email_subject)
        if parsed_email_subject == None:
            return True
        if isinstance(parsed_email_subject,list):
            for value in parsed_email_subject:
                for i in results.values():
                    if value.lower() == i['ticket'].lower():
                        return None
                    else:
                        continue
            return parsed_email_subject
            
        elif item['subject'].split() == parsed_email_subject.split():
            return None
    missing_sub_logger.warning("Email_Subject = {} : Parsed Subject = {}".format(email_subject,parsed_email_subject))
    return True

def split_subject(text):
    try:
        split_ticket = re.findall(ticket_reply_reg,text)
        ticket_no = re.findall(ticket_num_reg,text)
        if len(split_ticket):
            return ticket_no
        else:
            return text.strip()
    except Exception as e:
        error.exception(f'Subject_Split_Error: {e}')
        return None

if __name__=="__main__":
    pass