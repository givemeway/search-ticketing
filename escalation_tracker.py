from bs4 import BeautifulSoup
import bs4
import math
import csv
from itertools import zip_longest
from datetime import datetime
from searchTicket import search_query
from concurrent.futures import ThreadPoolExecutor
from parse import error
from copy import deepcopy
import os
import sys


time = datetime.fromtimestamp

n_workers = 5
idx = 0

csvFilePath = os.path.join(os.path.expanduser("~"), "Desktop", "escalated.csv")
agent_csv_file_path = os.path.join(
    os.path.expanduser("~"), "Desktop", "agents.csv")
departmentURL = "https://ticket.idrive.com/scp/ajax.php/tickets/1319774/transfer"
queryURL = 'https://ticket.idrive.com/scp/tickets.php?sort=date&dir=0&a=search&search-type=&query='

columns = ["URL", "Ticket", "Response to User", "Ticket Due", "Department", "Assigned", "Developer",
           "Developer Update", "Response to Dev", "Response Pending to Dev Update", "Notes", "Events"]

agent_tracker_cols = ["Ticket", "Thread_1",
                      "Thread_2", "Thread_3", "Thread_4", "Thread_5"]

exclude = ['India Billing Support', 'India Support', 'Indiasupport Supervisors', 'Indiasupport SPAM',
           'IndiaSupport Crisis', 'IDrive Support', 'IBackup Support', 'IndiaSupport Review',
           'RemotePC Support', 'Supervisors', 'US Support', 'US Development', 'Sales',
           'BMR Support', 'BMR Critcal', 'Infra Maintenance', 'BMR DevOps', 'Finance', 'BizDev',
           'Returns', 'Dev', 'Spam', 'Support', 'Artificial Intelligence', 'Design', 'Office IT', 'GDPR',
           '— Select —', 'Express', 'IndiaSupport 360'
           ]
indiasupportAgents = ['Sandeep Kumar G R', 'Santanu Chowdhury', 'Nandu Suresh', 'Ravi Chandra', 'Manikandan Support', 'Goutham MN', 'Nikhil Ugrankar', 'Rahul Raj', 'supreeth s', 'Johnson Prabu', 'Santhosh Support', 'jagadeesh k', 'Harshith D', 'Shanmugaraj M', 'Goutam Support', 'Sumit Sangwai', 'Praveen Shivappa', 'karthik ramamurthy', 'Arun Krishnan', 'Akhil Pedada', 'Deepak Reddy', 'prasanjit chatterjee', 'Chandan Support', 'rajkumar kb',
                      'Chetan G', 'Suraj Hekare', 'Vinil PV', 'Rakesh V', 'Jaspher Issac', 'Vikram Mutharapu', 'Krishna Kumar', 'Akshay V', 'Suraj Kumar', 'Melvin mathew', 'Arjun nk', 'akash waghamare', 'Sai Deep', 'Prabhu PT', 'Karthik Natarajan', 'Chethan Kumar', 'Aneesh Ram', 'Johnny Thomas', 'Sandeep  Shenoy', 'Harish Babu', 'Laxmon Philip', 'Krishna Menedhar', 'Satish Ningaiah', 'Udaya narayan', 'Santanu Chowdhury', 'Sarthak HS']


exclude = set(exclude)
selector1 = ".thread-entry.note .header b"
selector2 = ".thread-entry.note .header time"
selector4 = ".thread-entry.response .header b"
selector5 = ".thread-entry.response .header time"
selector7 = ".thread-entry.message .header time"
selector6 = ".thread-entry.message .header b"
selector3 = ".thread-event.action .faded.description:has(>b:first-child):has(>time):has(>strong)"
depart_head_selector = "div#content > :nth-child(3) >:nth-child(1) >:nth-child(1) >:nth-child(1) >:nth-child(1) >:nth-child(1) >:nth-child(3) >:nth-child(1)"
depart_name_selector = "div#content > :nth-child(3) >:nth-child(1) >:nth-child(1) >:nth-child(1) >:nth-child(1) >:nth-child(1) >:nth-child(3) >:nth-child(2)"
assigned_agent_selector = "div#content > :nth-child(5) >:nth-child(1) >:nth-child(1) >:nth-child(1) >:nth-child(1) >:nth-child(1) >:nth-child(1) >:nth-child(2)"


if sys.platform.startswith('win'):
    desktop_path = os.path.join(os.path.expanduser("~"), "Desktop")


def extract_ticket_url(soup, ticket):
    url = ""
    for tr in soup.select(".list tbody tr"):
        for td in tr.select("td:nth-child(2)"):
            if (td.get_text().strip() == ticket):
                url = "https://ticket.idrive.com" + \
                    td.find_all("a")[0]['href']
                break
    return url


def find_responses(soup, selector1, selector2):
    responses = []
    for b, time in zip_longest(soup.select(selector1), soup.select(selector2)):
        #         print(b.get_text().strip(),time.get_text().strip())
        responses.append({"agent": b.get_text().strip(),
                          "time": time.get_text().strip(),
                          "timestamp": datetime.strptime(time.get_text().strip().replace("\u202f", " "),
                                                         '%m/%d/%y %I:%M %p').timestamp()
                          })
    return responses


def find_messages(soup, selector1, selector2):
    messages = []
    for b, time in zip_longest(soup.select(selector1), soup.select(selector2)):
        #         print(b.get_text().strip(),time.get_text().strip())
        messages.append({"agent": b.get_text().strip(),
                         "time": time.get_text().strip(),
                         "timestamp": datetime.strptime(time.get_text().strip().replace("\u202f", " "),
                                                        '%m/%d/%y %I:%M %p').timestamp()
                         })
    return messages


def find_notes(soup, selector1, selector2):
    # .thread-entry.message .header  --> user response
    # .thread-entry.note .header --> notes
    # .thread-event.action .faded.description  --> thread action
    # .thread-entry.response .header     --> agent response
    notes = []
    for b, time in zip_longest(soup.select(selector1), soup.select(selector2)):
        #         print(b.get_text().strip(),time.get_text().strip())
        notes.append({"agent": b.get_text().strip(),
                      "time": time.get_text().strip(),
                      "timestamp": datetime.strptime(time.get_text().strip().replace("\u202f", " "),
                                                     '%m/%d/%y %I:%M %p').timestamp()
                      })
    return notes


def find_ticket_events(soup, selector):
    events = []
    for strong in soup.select(selector):
        if ("transferred this to" in strong.get_text().strip()):
            item = {"agent":   strong.select("b")[0].get_text().strip(),
                    "department":   strong.select("strong")[0].get_text().strip(),
                    "time":   strong.select("time")[0].get_text().strip(),
                    "timestamp":   datetime.strptime(strong.select("time")[0].get_text().strip().replace("\u202f", " "),
                                                     '%m/%d/%y %I:%M %p').timestamp()
                    }
            events.append(item)
    return events


def get_agents_or_departments(session, url):
    ticket = search_query(session, ticketURL=url)
    soup = BeautifulSoup(ticket.content, 'html5lib')
    dropdownitems = []
    for select in soup.select(".form-simple select"):
        for option in select.find_all("option"):
            department = option.get_text().strip().split(
                "Support /")[-1].strip()
            if department not in exclude:
                dropdownitems.append(department)
    return dropdownitems


def find_last_escalated_item(events, departments):
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
    columns = ['ticket', 'department']
    try:
        with open(csvFilePath, 'r', encoding='utf-8') as file:
            reader = csv.reader(file, delimiter=',')
            for row in reader:
                if row != columns:
                    img_obj[row[1]] = row[0]
        return img_obj
    except Exception as e:
        error.exception(e)
        return e


def extract_agents(session, departments, department, visited):
    try:
        agentsURL = departments[department]
        agents = set(get_agents_or_departments(session, agentsURL))
        if department == "India Support":
            agents = set(indiasupportAgents)
            visited["India Support"] = agents
        else:
            visited[department] = agents - set(indiasupportAgents)
        return agents

    except Exception as e:
        return None


def find_due_date(session, departments, escalated,
                  notes, ticket, visited, assigned_agent, responses):
    i = 0
    dev = None
    diff = None
    updateDuration = None
    agent_res = None
    agent_res_time = None
    dev_update_time = None
    agent_update_pending = None
    notes_len = len(notes)
    responses_len = len(responses)
    j = 0

    if "India Support" not in visited:
        extract_agents(session, departments, "India Support", visited)

    if escalated is not None:
        if escalated['department'] not in visited:
            extract_agents(session, departments,
                           escalated['department'], visited)
    while i < notes_len:
        note = notes.pop()

        if escalated is not None and note['timestamp'] > escalated['timestamp'] and \
                note['agent'] in visited[escalated['department']]:
            updateDuration = datetime.now() - time(note['timestamp'])
            dev_update_time = time(note['timestamp'])
            j = 0
            while j < responses_len:
                res = responses.pop()
                if res['timestamp'] > note['timestamp']:
                    if (res['agent'] == assigned_agent or res['agent'] in visited['India Support']):
                        agent_res = res
                else:
                    break
                j += 1

            if agent_res is not None:
                agent_res_time = time(
                    agent_res['timestamp']) - time(note['timestamp'])
            dev = note
            break

        i += 1

    if updateDuration is None and escalated is not None:
        diff = datetime.now() - time(escalated['timestamp'])
    if updateDuration is not None and agent_res is None:
        agent_update_pending = datetime.now() - dev_update_time

    return {"ticket": ticket,
            "due": str(round(diff.total_seconds() / 86400)) if diff is not None else str(None),
            "escalated": escalated,
            "Developer Update": str(updateDuration),
            "updatedAgent": dev,
            "assignedAgentUpdate": str(agent_res_time) if agent_res_time else None,
            "agentUpdatePending": str(agent_update_pending) if agent_update_pending else None
            }


def get_ticket_html_content(session, ticket):
    try:
        ticketSearch = search_query(session, searchQuery=ticket)
        searchSoup = BeautifulSoup(ticketSearch.content, 'html5lib')
        ticketURL = extract_ticket_url(searchSoup, ticket)
        ticketContent = search_query(session, ticketURL=ticketURL)
        soup = BeautifulSoup(ticketContent.content, 'html5lib')
        assigned_department = soup.select(depart_name_selector)[
            0].get_text().split("Support /")[-1].strip()
        assigned_agent = soup.select(assigned_agent_selector)[
            0].get_text().split("Support /")[-1].strip()
        return (soup, assigned_department, assigned_agent, ticketURL)
    except Exception as e:
        error.exception(e)
        return None


def find_res_pending(responses, messages):
    i = 0
    res_len = len(responses)
    no_response = None
    if messages:
        msg = messages.pop()
        while i < res_len:
            res = responses.pop()
            if msg['timestamp'] > res['timestamp']:
                no_response = msg
            else:
                break
            i += 1
        return str(datetime.now() - time(no_response['timestamp'])) if no_response else "Updated"
    return "Updated"


def process_ticket(ticket, visited, payload):
    global idx
    idx += 1
    payload['processing'].emit((ticket, idx))
    session = payload['session']
    departments = payload['departments']
    agents_in_departments = payload['agents_in_department']
    html_soup = get_ticket_html_content(session, ticket)
    if html_soup is not None:
        soup, assigned_depart, assigned_agent, ticketURL = html_soup
        if (isinstance(soup, bs4.BeautifulSoup)):
            notes = find_notes(soup, selector1, selector2)
            responses = find_responses(soup, selector4, selector5)
            messages = find_messages(soup, selector6, selector7)
            deepcopy_res = deepcopy(responses)
            events = find_ticket_events(soup, selector3)
            notes_len = len(notes)
            events_len = len(events)
            escalated = find_last_escalated_item(events, departments)
            due = find_due_date(session, agents_in_departments,
                                escalated, notes, ticket, visited,
                                assigned_agent, responses)
            due['UsersResponsePending'] = find_res_pending(
                deepcopy_res, messages)
            due['notes'] = notes_len
            due['events'] = events_len
            due['department'] = assigned_depart
            due['assigned'] = assigned_agent
            due['computed'] = True
            due['url'] = ticketURL
            per = math.ceil((idx/payload['total'])*100)
            payload["pBar"].emit(per)
            return due
        else:
            return {"ticket": ticket, "computed": False}
    else:
        return {"ticket": ticket, "computed": False}


def escalation_worker(_dict):
    global idx
    processed = []
    idx = 0
    _dict['agents_in_department'] = extract_departments(
        "./logs/departments.csv")
    visited = {}
    session = _dict['session']
    _dict['departments'] = set(
        get_agents_or_departments(session, departmentURL))
    with ThreadPoolExecutor(n_workers) as exe:
        futures = [exe.submit(process_ticket, ticket, visited, _dict)
                   for ticket in _dict['tickets']]
        for future in futures:
            due = future.result()
            processed.append(due)

    df = create_dataframe_csv(processed)
    write_csv(csvFilePath, columns, df)


def msg_res_eta(ticket, responses, messages):
    time_diff_dict = {"ticket": ticket, "computed": True}
    if messages:
        for idx in range(5):
            if messages:
                if len(messages) > 0 and len(responses) > 0:
                    msg = messages.pop()
                    res = responses.pop()
                    time_diff = res['timestamp'] - msg['timestamp']
                    temp_main_diff = str(
                        time(res['timestamp'])-time(msg['timestamp']))
                    time_diff_dict["Thread_"+str(idx+1)] = "Received: " + msg['time'] + \
                        " || " + "Replied: " + temp_main_diff
                    if (time_diff > 3600):
                        while len(responses) > 0:
                            res = responses.pop()
                            time_diff = res['timestamp'] - msg['timestamp']
                            temp_diff = str(
                                time(res['timestamp'])-time(msg['timestamp']))
                            if time_diff > 3600:
                                time_diff_dict["Thread_"+str(idx+1)] = "Received: " + \
                                    msg['time'] + " || " + \
                                    "Replied: " + temp_diff
                                continue
                            elif time_diff < 0:
                                responses.append(res)
                                break
                            elif time_diff <= 3600:
                                time_diff_dict["Thread_"+str(idx+1)] = "Received: " + \
                                    msg['time'] + " || " + \
                                    "Replied: " + temp_diff
                                break
                    elif time_diff < 0:
                        responses.append(res)
    return time_diff_dict


def process_agent_ticket(ticket, visited, payload):
    global idx
    idx += 1
    payload['processing'].emit((ticket, idx))
    session = payload['session']
    html_soup = get_ticket_html_content(session, ticket)
    if html_soup is not None:
        soup, assigned_depart, assigned_agent, ticketURL = html_soup
        if (isinstance(soup, bs4.BeautifulSoup)):
            responses = find_responses(soup, selector4, selector5)
            messages = find_messages(soup, selector6, selector7)
            deepcopy_res = deepcopy(responses)
            res_mapped = msg_res_eta(
                ticket, deepcopy_res, messages)
            per = math.ceil((idx/payload['total'])*100)
            payload["pBar"].emit(per)
            return res_mapped
        else:
            return {"ticket": ticket, "computed": False}
    else:
        return {"ticket": ticket, "computed": False}


def agent_tracker_worker(_dict):
    global idx
    processed = []
    idx = 0
    _dict['agents_in_department'] = extract_departments(
        "./logs/departments.csv")
    visited = {}
    session = _dict['session']
    _dict['departments'] = set(
        get_agents_or_departments(session, departmentURL))

    with ThreadPoolExecutor(n_workers) as exe:
        futures = [exe.submit(process_agent_ticket, ticket, visited, _dict)
                   for ticket in _dict['tickets']]
        for future in futures:
            due = future.result()
            processed.append(due)
    df = create_agent_dataframe_csv(processed)
    write_csv(agent_csv_file_path, agent_tracker_cols, df)


def create_agent_dataframe_csv(processed_tickets):
    df = []
    for row in processed_tickets:
        if row['computed']:
            df.append([row['ticket'],
                       row["Thread_1"] if "Thread_1" in row else "",
                       row["Thread_2"] if "Thread_2" in row else "",
                       row["Thread_3"] if "Thread_3" in row else "",
                       row["Thread_4"] if "Thread_4" in row else "",
                       row["Thread_5"] if "Thread_5" in row else "",
                       ])
        else:
            df.append([row['ticket'],
                       "Compute Error",
                       "Compute Error",
                       "Compute Error",
                       "Compute Error",
                       "Compute Error",
                       ])
    return df


def create_dataframe_csv(processed_tickets):
    df = []
    for row in processed_tickets:
        if row['computed']:
            df.append([
                row['url'],
                row['ticket'],
                row['UsersResponsePending'],
                row["due"],
                row['department'],
                row['assigned'],
                None if row["updatedAgent"] is None else row["updatedAgent"]["agent"],
                None if row["updatedAgent"] is None else
                datetime.now() - time(row["updatedAgent"]['timestamp']),
                row['assignedAgentUpdate'],
                row["agentUpdatePending"],
                row['notes'],
                row['events']
            ])
        else:
            df.append([
                "Compute Error",
                row['ticket'],
                "Compute Error",
                "Compute Error",
                "Compute Error",
                "Compute Error",
                "Compute Error",
                "Compute Error",
                "Compute Error",
                "Compute Error",
                "Compute Error",
                "Compute Error"
            ])
    return df


def write_csv(csvFilePath, columns, arr):
    try:
        with open(csvFilePath, 'w', newline='', encoding='utf-8') as file:
            writer = csv.writer(file, delimiter=',')
            writer.writerow(columns)
            for row in arr:
                writer.writerow(row)
    except Exception as e:
        error.exception(e)
        return e


if __name__ == "__main__":
    pass
