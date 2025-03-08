import os
import csv
import bs4
from urllib.parse import urlparse, parse_qs
from escalation_tracker import get_ticket_html_content, find_notes

selector1 = ".thread-entry.note .header b"
selector2 = ".thread-entry.note .header time"

output_csv = os.path.join(os.path.expanduser("~"), "Desktop", "audit.csv")


def get_file_info(session, file_path):
    try:
        file_name = os.path.basename(file_path)
        file_size = os.path.getsize(file_path) / 1024  # Size in bytes

        with open(file_path, 'r', encoding='utf-8') as file:
            lines = file.readlines()
            line_count = len(lines)
        ticket = file_name.split(".")[0].split("-")[1].strip()
        # html_soup = get_ticket_html_content(session, ticket)
        # if html_soup is not None:
        #     soup, _, _, url = html_soup
        #     notes = find_notes(soup, selector1, selector2)

        return {
            "file_name": file_name.split(".")[0],
            "ticket": ticket,
            "line_count": line_count,
            "file_size": '{}kb'.format(round(file_size, 2)),
            "size": round(file_size, 2)
        }
    except Exception as e:
        return {"error": str(e)}


def call_audit_worker(session, directory_path):
    file_data = []

    for file_name in os.listdir(directory_path):
        file_path = os.path.join(directory_path, file_name)
        if os.path.isfile(file_path) and file_name.endswith(".txt"):
            info = get_file_info(session, file_path)
            if "error" not in info:
                file_data.append([info["file_name"], info["line_count"],
                                 info["file_size"], info["size"], info["ticket"]])

    file_data.sort(key=lambda x: x[3], reverse=True)

    with open(output_csv, mode="w", newline="", encoding="utf-8") as csv_file:
        writer = csv.writer(csv_file)
        writer.writerow(["filename", "line_count",
                        "size", "file_size", "ticket"])
        writer.writerows(file_data)

    print(f"File info saved to {output_csv}")


if __name__ == "__main__":
    # Replace with your directory path
    pass
