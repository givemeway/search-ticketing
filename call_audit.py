import os
import csv
import bs4
from urllib.parse import urlparse, parse_qs
from escalation_tracker import get_ticket_html_content, find_notes_with_body, extract_agents, extract_departments
from concurrent.futures import ThreadPoolExecutor
import math
selector1 = ".thread-entry.note .header b"
selector2 = ".thread-entry.note .header time"

output_csv = os.path.join(os.path.expanduser("~"), "Desktop", "audit.csv")
audit_note_csv = os.path.join(
    os.path.expanduser("~"), "Desktop", "audit_notes.csv")
note_threads = 3
idx = 0


def get_file_info(session, file_path):
    try:
        file_name = os.path.basename(file_path)
        file_size = os.path.getsize(file_path) / 1024  # Size in bytes

        with open(file_path, 'r', encoding='utf-8') as file:
            lines = file.readlines()
            line_count = len(lines)
        ticket = file_name.split(".")[0].split("-")[1].strip()
        agent_name = file_name.split(".")[0].split("-")[0].strip()
        # html_soup = get_ticket_html_content(session, ticket)
        # if html_soup is not None:
        #     soup, _, _, url = html_soup
        #     notes = find_notes(soup, selector1, selector2)

        return {
            "agent_name": agent_name,
            "file_name": file_name.split(".")[0],
            "ticket": ticket,
            "line_count": line_count,
            "file_size": '{}kb'.format(round(file_size, 2)),
            "size": round(file_size, 2)
        }
    except Exception as e:
        return {"error": str(e)}


def audit_notes(_dict):
    departments = extract_departments(
        "./logs/departments.csv")
    visited = {}
    agents = extract_agents(
        _dict["session"], departments, "India Support", visited)
    print("Agents _ main: ", agents)
    processed = []
    with ThreadPoolExecutor(5) as executor:
        futures = [executor.submit(
            audit_notes_worker, _dict["session"], ticket, agents, _dict) for ticket in _dict["tickets"]]
        for future in futures:
            processed.append(future.result())
    df = create_dataframe_csv(processed)
    columns = [
        "agent_name", "ticket", "line_count", "size"]
    columns.extend(["body_{}".format(i) for i in range(1, note_threads + 1)])
    print("Columns: ", columns)
    write_csv(audit_note_csv, columns, df)
    print(f"Audit notes saved to {audit_note_csv}")


def create_dataframe_csv(processed_tickets):
    df = []

    for row in processed_tickets:
        if row is not None:
            cols = [row["agent_name"], row["ticket"], row["line_count"], row["size"]
                    ]
            cols.extend([row["body_{}".format(i)]
                        for i in range(1, note_threads + 1)])
            df.append(cols)
    return df


def write_csv(pth, columns, df):
    try:
        with open(pth, "w", newline="", encoding="utf-8") as file:
            writer = csv.writer(file)
            writer.writerow(columns)
            writer.writerows(df)
    except Exception as e:
        print(e)


def audit_notes_worker(session, ticket, agents, payload):
    global idx
    idx += 1
    payload['processing'].emit((ticket[0], idx))
    html_soup = get_ticket_html_content(session, ticket[5])
    extracted_note = {"agent_name": ticket[0], "file_name": ticket[1],
                      "line_count": ticket[2], "size": ticket[3], "ticket": ticket[5]}
    for i in range(1, note_threads + 1):
        extracted_note["body_{}".format(i)] = ""
    if html_soup is not None:
        soup, _, _, url = html_soup
        notes = find_notes_with_body(soup, selector1, selector2)
        for count in range(note_threads):
            extracted_note['body_{}'.format(count+1)] = ""
            if len(notes) > 0:
                note = notes.pop()
                search_term = ticket[0].lower()
                matches = []
                for name in agents:
                    if search_term.lower() in name.lower():
                        print("Matched: ", name)
                        matches.append(name)
                agent_full_name = ""
                if len(matches) == 1:
                    agent_full_name = matches[0]
                if len(matches) > 1:
                    agent_full_name = matches[-1]
                print("Agent full name", agent_full_name)
                if agent_full_name == note['agent']:
                    extracted_note['body_{}'.format(count+1)] = note['body']
        per = math.ceil((idx/payload['total'])*100)
        payload["pBar"].emit(per)
    return extracted_note


def call_audit_worker(session, directory_path):
    file_data = []

    for file_name in os.listdir(directory_path):
        file_path = os.path.join(directory_path, file_name)
        if os.path.isfile(file_path) and file_name.endswith(".txt"):
            info = get_file_info(session, file_path)
            if "error" not in info:
                file_data.append([info["agent_name"], info["file_name"], info["line_count"],
                                 info["file_size"], info["size"], info["ticket"]])

    file_data.sort(key=lambda x: x[3], reverse=True)

    with open(output_csv, mode="w", newline="", encoding="utf-8") as csv_file:
        writer = csv.writer(csv_file)
        writer.writerow(["agent_name", "filename", "line_count",
                        "size", "file_size", "ticket"])
        writer.writerows(file_data)

    print(f"File info saved to {output_csv}")


if __name__ == "__main__":
    # Replace with your directory path
    pass
