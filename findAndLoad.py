from searchTicket import search_query
from login import create_connection
from bs4 import BeautifulSoup
from searchTicket import idriveinc_domain_regex
import re

def extract_search(soup):
    list_obj ={}
    link ="https://ticket.idrive.com/scp/"
    for count,tr in enumerate(soup.find_all('tr')):
        search_obj={}
        for i,td in enumerate(tr.find_all('td')):
            if i==1:
                search_obj['ticket'] ={}
                search_obj['ticket'][0] = td.text.strip()
                for a in enumerate(td.find_all('a',href=True)):
                    search_obj['ticket'][1] = link + a[1]['href']
            elif i==2:
                search_obj['created']= td.text.strip()
            elif i==3:
                children = td.findChildren('div',recursive=True)
                attachments = td.findChildren('i',attrs={"class":"icon-paperclip"},recursive=True)
                for child in children:
                    search_obj['subject'] = child.text.strip()
                if len(attachments):
                    for _ in attachments:
                        search_obj['attachment'] = 1
            elif i==5:
                search_obj['status']= td.text.strip()
            else:
                continue
        if len(search_obj):       
            list_obj[count] = search_obj
        
    return list_obj

def search_ticket(s,string,result):
    r = search_query(s,string)
    soup = BeautifulSoup(r.content,'html5lib')
    data = extract_search(soup)
    idx = 1
    if len(data) > 0:
        for item in data.values():
            result.emit([idx,item['ticket'],item['subject'],item['status'],item['created'],item['attachment'] if 'attachment' in item else 0])
            idx +=1
    else:
        result.emit(None)
def loadPrevious(id,email,subject):
                conn = create_connection('ticket.db')
                if conn is not None:
                    with conn:
                        cur = conn.cursor()
                        cur.execute('''SELECT * FROM mbox WHERE ID=?''',(int(id),))
                        row = cur.fetchone()
                        if row is not None:
                            start = row[3]
                            end = row[4]
                            try:
                                email_start = row[5]+1
                                subject_start = row[7]+1
                            except Exception as e:
                                email_start = row[5]
                                subject_start = row[7]
                            email_end = row[6]
                            subject_end = row[8]
                            cur.execute('''SELECT * FROM emails WHERE ID BETWEEN ? and ?''',(start,end))
                            all_emails = cur.fetchall()
                            cur.execute('''SELECT * FROM emailnotfound WHERE ID BETWEEN ? and ?''',(email_start,email_end))
                            emails_not_found = cur.fetchall()
                            cur.execute('''SELECT * FROM subjectnotfound WHERE ID BETWEEN ? and ?''',(subject_start,subject_end))
                            subject_not_found = cur.fetchall()
                            for i, item in enumerate(emails_not_found):
                                if len(re.findall(idriveinc_domain_regex,item[1]))==0:
                                    email.emit([i+1,item])
                            for i, item in enumerate(subject_not_found):
                                if len(re.findall(idriveinc_domain_regex,item[1]))==0:
                                    subject.emit([i+1,item])