from bs4 import BeautifulSoup
import bs4
import math,csv
from urllib.parse import urlparse, parse_qs
from itertools import zip_longest
from datetime import datetime
from searchTicket import search_query
from concurrent.futures import ThreadPoolExecutor
from parse import error

n_workers = 5
idx = 0

csvFilePath ="./escalated.csv"
departmentURL = "https://ticket.idrive.com/scp/ajax.php/tickets/1319774/transfer"
queryURL ='https://ticket.idrive.com/scp/tickets.php?sort=date&dir=0&a=search&search-type=&query='

columns = ["Ticket","Due Since","Department","Assigned Agent","Developer","LastUpdated","Notes","Events"]

exclude = ['India Billing Support','India Support','Indiasupport Supervisors','Indiasupport SPAM',\
           'IndiaSupport Crisis','IDrive Support','IBackup Support','IndiaSupport Review',\
           'RemotePC Support','Supervisors','US Support','US Development','Sales',\
           'BMR Support','BMR Critcal','Infra Maintenance','BMR DevOps','Finance','BizDev',\
           'Returns','Dev','Spam','Support','Artificial Intelligence','Design','Office IT','GDPR',\
           '— Select —','Express','IndiaSupport 360'
          ]
exclude = set(exclude)
selector1 = ".thread-entry.note .header b"
selector2 = ".thread-entry.note .header time"
selector3 = ".thread-event.action .faded.description:has(>b:first-child):has(>time):has(>strong)"
depart_head_selector = "div#content > :nth-child(3) >:nth-child(1) >:nth-child(1) >:nth-child(1) >:nth-child(1) >:nth-child(1) >:nth-child(3) >:nth-child(1)"
depart_name_selector = "div#content > :nth-child(3) >:nth-child(1) >:nth-child(1) >:nth-child(1) >:nth-child(1) >:nth-child(1) >:nth-child(3) >:nth-child(2)"
assigned_agent_selector = "div#content > :nth-child(5) >:nth-child(1) >:nth-child(1) >:nth-child(1) >:nth-child(1) >:nth-child(1) >:nth-child(1) >:nth-child(2)"


def extract_ticket_url(soup,ticket):
    url = ""
    for tr in soup.select(".list tbody tr"):
        for td in tr.select("td:nth-child(2)"):
            if(td.get_text().strip() == ticket):
                url = "https://ticket.idrive.com/scp/" + td.find_all("a")[0]['href']
                break
    return url

def find_notes(soup,selector1,selector2):
    # .thread-entry.message .header  --> user response
    # .thread-entry.note .header --> notes
    # .thread-event.action .faded.description  --> thread action    
    # .thread-entry.response .header     --> agent response
    notes = []
    for b,time in zip_longest(soup.select(selector1),soup.select(selector2)):
#         print(b.get_text().strip(),time.get_text().strip())
        notes.append({"agent":b.get_text().strip(),\
                      "time":time.get_text().strip(),\
                      "timestamp": datetime.strptime(time.get_text().strip(),\
                                 '%m/%d/%y, %I:%M %p').timestamp()    
                     })
    return notes

def find_ticket_events(soup,selector):
    events = []
    for strong in soup.select(selector):
        if("transferred this to" in strong.get_text().strip()):
            item = {  "agent"     :   strong.select("b")[0].get_text().strip(),\
                      "department":   strong.select("strong")[0].get_text().strip(),\
                      "time"      :   strong.select("time")[0].get_text().strip(),\
                      "timestamp" :   datetime.strptime(strong.select("time")[0].get_text().strip(),\
                                     '%m/%d/%y, %I:%M %p').timestamp()  
                    }
            events.append(item)
    return events

def get_agents_or_departments(session,url):
    ticket = search_query(session,url)
    soup = BeautifulSoup(ticket.content,'html5lib')
    dropdownitems = []
    for select in soup.select(".form-simple select"):
        for option in select.find_all("option"):
            department = option.get_text().strip().split("Support /")[-1].strip()
            if department not in exclude:
                dropdownitems.append(department)
    return dropdownitems

def find_last_escalated_item(events,departments):
    i = 0
    escalated = None
    events_len = len(events)
    while i < events_len:
        item = events.pop()
        if item["department"] in departments:
            escalated = item
            break
        i += 1
    return escalated

def extract_departments(csvFilePath):
    img_obj = {}
    columns = ['ticket','department']
    try:
        with open(csvFilePath,'r',encoding='utf-8') as file:
            reader = csv.reader(file,delimiter=',')
            for row in reader:
                if row != columns:
                    img_obj[row[1]] = row[0]
        return img_obj
    except Exception as e:
        error.exception(e)
        return e

def extract_agents(session,departments,department,visited):
    ticket = departments[department]
    ticketSearch = search_query(session,searchQuery=ticket)
    searchSoup = BeautifulSoup(ticketSearch.content,'html5lib')
    ticketURL = extract_ticket_url(searchSoup,ticket)
    parse_result = urlparse(ticketURL)
    ticketID = parse_qs(parse_result.query)['id'][0]
    agentsURL = "https://ticket.idrive.com/scp/ajax.php/tickets/" + ticketID + "/assign/agents"
    agents = set(get_agents_or_departments(session,agentsURL))
    visited[department] = agents
    return agents
    
def find_due_date(session,departments,escalated,notes,ticket,visited):
    i = 0
    dev = None
    diff = None
    updateDuration = None
    notes_len = len(notes)
    if escalated is not None:
        if escalated['department'] not in visited:
             extract_agents(session,departments,escalated['department'],visited)
    while i < notes_len:
        item = notes.pop()
        if escalated is not None and item['timestamp'] > escalated['timestamp'] and \
            item['agent'] in visited[escalated['department']]:
                
            updateDuration = datetime.fromtimestamp(item['timestamp']) - \
            datetime.fromtimestamp(escalated['timestamp'])
            dev = item
            break
        i += 1
    if updateDuration is None and escalated is not None:
        diff = datetime.now() - datetime.fromtimestamp(escalated['timestamp'])
        
    return {"ticket":ticket, \
            "due":str(diff),\
            "escalated":escalated,\
            "ifUpdatedTimeTaken":str(updateDuration),\
            "updatedAgent":dev
           }


def get_ticket_html_content(session,ticket):
    try:
        ticketSearch = search_query(session, searchQuery = ticket)
        searchSoup = BeautifulSoup(ticketSearch.content,'html5lib')
        ticketURL = extract_ticket_url(searchSoup,ticket)
        ticketContent = search_query(session,ticketURL)
        soup = BeautifulSoup(ticketContent.content,'html5lib')
        assigned_department = soup.select(depart_name_selector)[0].get_text().split("Support /")[-1].strip()
        assigned_agent = soup.select(assigned_agent_selector)[0].get_text().split("Support /")[-1].strip()
        return (soup,assigned_department,assigned_agent)
    except Exception as e:
        error.exception(e)
        return None
    
def process_ticket(ticket,visited,payload):
    global idx
    idx += 1
    payload['processing'].emit((ticket,idx))
    session = payload['session']
    departments = payload['departments']
    agents_in_departments = payload['agents_in_department']
    soup,assigned_depart,assigned_agent = get_ticket_html_content(session,ticket)
    if(isinstance(soup,bs4.BeautifulSoup)):
        notes = find_notes(soup,selector1,selector2)
        notes_len = len(notes)
        events = find_ticket_events(soup,selector3)
        events_len = len(events)
        escalated = find_last_escalated_item(events,departments)
        due = find_due_date(session,agents_in_departments,escalated,notes,ticket,visited)
        due['notes'] = notes_len
        due['events'] = events_len
        due['department'] = assigned_depart
        due['assigned'] = assigned_agent
        due['computed'] = True
        per = math.ceil( (idx/payload['total'])*100 ) 
        payload["pBar"].emit(per)
        return due
    else:
        return {"ticket":ticket,"computed":False}

def escalation_worker(_dict):
    processed = []
    _dict['agents_in_department'] = extract_departments("./departments.csv")
    visited = {}
    session = _dict['session']
    _dict['departments'] = set(get_agents_or_departments(session,departmentURL))

    with ThreadPoolExecutor(n_workers) as exe:
        futures = [exe.submit(process_ticket,ticket,visited,_dict) for ticket in _dict['tickets']]
        for future in futures:
            due = future.result()
            processed.append(due)

    df = create_dataframe_csv(processed)
    write_csv(csvFilePath,columns,df)
    
def create_dataframe_csv(processed_tickets):
    df = []
    for row in processed_tickets:
        if row['computed']:
            df.append([row['ticket'],\
                            row["due"],\
                            row['department'],\
                            row['assigned'],\
                            None if row["updatedAgent"] is None else row["updatedAgent"]["agent"],\
                            None if row["updatedAgent"] is None else \
                            datetime.now() - datetime.fromtimestamp(row["updatedAgent"]['timestamp']),\
                            row['notes'],\
                            row['events']
                        ])
        else:
            df.append([row['ticket'],\
                            "Compute Error",\
                            "Compute Error",\
                            "Compute Error",\
                            "Compute Error",\
                            "Compute Error",\
                            "Compute Error",\
                            "Compute Error"
                        ])
    return df

def write_csv(csvFilePath,columns,arr):
    try:
        with open(csvFilePath,'w',newline='',encoding='utf-8') as file:
            writer = csv.writer(file,delimiter=',')
            writer.writerow(columns)
            for row in arr:
                writer.writerow(row)
    except Exception as e:
        error.exception(e)
        return e

if __name__=="__main__":
    pass