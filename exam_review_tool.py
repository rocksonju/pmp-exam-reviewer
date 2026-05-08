import argparse
import re
import pandas as pd
from pathlib import Path

def parse_markdown(md_path: Path):
    """Parse the custom exam_input markdown format.

    Extracts fields:
    - QuestionNo, QID
    - Question (description before choices)
    - Choices (concatenated)
    - SelectedChoice
    - ProcessGroup
    - TaskArea
    - Justification
    - Domain, Task, Reference (if present)
    """
    content = md_path.read_text(encoding="utf-8")
    # Split blocks by Question No. to keep multiline questions intact
    blocks = re.split(r"(?=^Question No\.)", content, flags=re.MULTILINE)
    results = []
    
    for block in blocks:
        block = block.strip()
        if not block:
            continue
        lines = block.splitlines()
        
        q_no = ""
        q_id = ""
        question_text = []
        choices = []
        selected = []
        process_group = []
        task_area = []
        justification = []
        domain = []
        task_domain = []
        reference = []
        
        state = "question"
        
        for line in lines:
            line_str = line.strip()
            if not line_str and state not in ["question", "justification"]:
                continue
                
            m = re.match(r"^Question No\.\s*(.+?)\s*-\s*ID\s*:\s*(.+)", line_str)
            if m:
                q_no = m.group(1).strip()
                q_id = m.group(2).strip()
                state = "question"
                continue
                
            if re.match(r"^Choice\s*\d+", line_str):
                state = "choice"
                continue
                
            if line_str == "Selected Choice:":
                state = "selected"
                continue
            if line_str == "Process Group :":
                state = "process"
                continue
            if line_str == "Task/K. Area :":
                state = "task_area"
                continue
            if line_str == "Justification:":
                state = "justification"
                continue
                
            if line_str.startswith("Domain:"):
                state = "domain"
                domain.append(line_str.split(":", 1)[1].strip())
                continue
            if line_str.startswith("Task:"):
                state = "task_domain"
                task_domain.append(line_str.split(":", 1)[1].strip())
                continue
            if line_str.startswith("Reference:"):
                state = "reference"
                reference.append(line_str.split(":", 1)[1].strip())
                continue
                
            if state == "choice":
                choices.append(line_str)
                state = "await"
            elif state == "selected":
                selected.append(line_str)
            elif state == "process":
                process_group.append(line_str)
            elif state == "task_area":
                task_area.append(line_str)
            elif state == "justification":
                justification.append(line)
            elif state == "domain":
                domain.append(line_str)
            elif state == "task_domain":
                task_domain.append(line_str)
            elif state == "reference":
                reference.append(line_str)
            elif state == "question":
                question_text.append(line)

        results.append({
            "QuestionNo": q_no,
            "QID": q_id,
            "Question": "\n".join(question_text).strip(),
            "Choices": "; ".join([c for c in choices if c]),
            "SelectedChoice": " ".join(selected).strip(),
            "ProcessGroup": " ".join(process_group).strip(),
            "TaskArea": " ".join(task_area).strip(),
            "Justification": "\n".join(justification).strip(),
            "Domain": " ".join(domain).strip(),
            "Task": " ".join(task_domain).strip(),
            "Reference": " ".join(reference).strip(),
        })
    return results

def export_to_excel(data, excel_path: Path):
    """Export parsed data to an Excel file with detailed columns."""
    df = pd.DataFrame(data)
    cols = [
        "QuestionNo", "QID", "Question", "Choices", "SelectedChoice",
        "ProcessGroup", "TaskArea", "Justification", "Domain", "Task", "Reference"
    ]
    df = df[cols]
    df.to_excel(excel_path, index=False)
    print(f"Excel file written to {excel_path}")


def export_to_markdown(data, md_path: Path):
    """Export parsed data to a markdown file matching the desired output format."""
    lines = []
    for entry in data:
        lines.append(f"Question No. :{entry.get('QuestionNo','')}")
        lines.append(f"Q - ID : {entry.get('QID','')}")
        lines.append("")
        lines.append("Question: ")
        lines.append("")
        lines.append(f"{entry.get('Question','')}")
        lines.append("")
        lines.append("Answer: options")
        
        choices = entry.get('Choices','').split('; ')
        for idx, choice in enumerate(choices, start=1):
            lines.append(f"Choice {idx} : {choice}")
        lines.append("")
        lines.append(f"Selected Choice: {entry.get('SelectedChoice','')}")
        lines.append("")
        lines.append(f"Process Group : {entry.get('ProcessGroup','')}")
        lines.append("")
        lines.append(f"Task/K. Area : {entry.get('TaskArea','')}")
        lines.append("")
        
        just_text = entry.get('Justification','').strip().lstrip('"').rstrip('"')
        ref_text = entry.get('Reference','').strip().rstrip(';"').rstrip()
        
        lines.append("Justification:")
        lines.append(f"\"{just_text}")
        lines.append("")
        if entry.get('Domain'):
            lines.append(f"Domain: {entry.get('Domain','')};")
        if entry.get('Task'):
            lines.append(f"Task: {entry.get('Task','')};")
        if entry.get('Reference'):
            lines.append(f"Reference: {ref_text} ;\"")
        lines.append("")
        lines.append("")
    md_path.write_text("\n".join(lines), encoding="utf-8")
    print(f"Markdown file written to {md_path}")

def main():
    parser = argparse.ArgumentParser(description="Convert custom exam markdown to Excel and Markdown")
    parser.add_argument("markdown", type=Path, help="Path to the source .md file")
    parser.add_argument("excel", type=Path, nargs="?", default=None,
                        help="Path to output .xlsx file (optional). If omitted, uses same name as markdown with .xlsx extension.")
    parser.add_argument("mdout", type=Path, nargs="?", default=None,
                        help="Path to output .md file (optional). If omitted, uses same name as markdown with .out.md extension.")
    args = parser.parse_args()
    md_path = args.markdown
    if not md_path.is_file():
        raise FileNotFoundError(f"Markdown file not found: {md_path}")
    excel_path = args.excel or md_path.with_suffix('.xlsx')
    md_out_path = args.mdout or md_path.with_name(md_path.stem + '.out.md')
    data = parse_markdown(md_path)
    export_to_excel(data, excel_path)
    export_to_markdown(data, md_out_path)

if __name__ == "__main__":
    main()
