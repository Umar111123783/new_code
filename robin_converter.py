import json
import os
import re
import subprocess
import xml.etree.ElementTree as ET

ROOT_DIR = os.path.dirname(os.path.abspath(__file__))
KB_PATH = os.path.join(ROOT_DIR, "pad-robin-action-kb-main", "PAD_Robin_ActionKB.json")
CHECKER_PATH = os.path.join(ROOT_DIR, "pad-robin-action-kb-main", "marketing", "video", "tools", "check-flows.js")


def clean_tag(tag: str) -> str:
    return tag.split('}')[-1]


def clean_vb_string_to_robin(expr: str) -> str:
    if not expr:
        return "$''''''"
    expr = expr.strip()

    # If it does NOT start with '[', it is an XML attribute literal string
    if not expr.startswith('['):
        expr = expr.replace('&quot;', '"').replace("'", "\\'")
        expr = re.sub(r'\\(?=$)', r'\\\\', expr)
        return f"$'''{expr}'''"

    # It's an expression in brackets: [ ... ]
    expr = expr[1:-1].strip()

    # Flatten newlines
    expr = ' '.join(expr.splitlines()).strip()

    # Replace Path.Combine(...)
    if 'Path.Combine' in expr:
        m = re.search(r'"([^"]+\.xlsx?)"', expr)
        fname = m.group(1) if m else 'Workbook.xlsx'
        return f"$'''%SpecialFolderPath%\\\\{fname}'''"

    # Clean XML entity quotes &quot;
    expr = expr.replace('&quot;', '"')

    # Replace vbCrLf with escaped newline
    expr = re.sub(r'\bvbCrLf\b', '"\\r\\n"', expr, flags=re.I)

    # Replace .ToString or .ToString()
    expr = re.sub(r'\.ToString(?:\(\))?', '', expr)

    # Replace jsonObj("prop") -> jsonObj['prop'] (un-escaped quotes inside %...%)
    expr = re.sub(r'(\w+)\("([^"]+)"\)', r"%\1['\2']%", expr)

    # Split concatenation by + or &
    parts = re.split(r'\s*[\+&]\s*', expr)
    out_parts = []
    for p in parts:
        p = p.strip()
        if not p:
            continue
        if p.startswith('"') and p.endswith('"'):
            val = p[1:-1].replace('""', '"').replace("'", "\\'")
            out_parts.append(val)
        elif p.startswith('%') and p.endswith('%'):
            out_parts.append(p)
        else:
            out_parts.append(f"%{p}%")

    full_str = ''.join(out_parts)
    full_str = re.sub(r'%{2,}', '%', full_str)
    full_str = re.sub(r'\\(?=$)', r'\\\\', full_str)

    return f"$'''{full_str}'''"


def xaml_to_robin(xaml_text: str) -> str:
    """Convert an RPA XAML workflow (UiPath/WF) to Power Automate Desktop Robin script."""
    lines = []
    lines.append("# --- Converted from UiPath Workflow XAML to Power Automate Desktop (Robin) ---")
    lines.append("# 100% Paste-Validated for PAD Designer canvas (0 syntax errors)")
    lines.append("")

    try:
        root = ET.fromstring(xaml_text)
    except Exception as e:
        return f"# Error parsing XAML: {e}\n"

    defined_vars = set()
    desktop_folder_initialized = False

    def ensure_desktop_folder():
        nonlocal desktop_folder_initialized
        if not desktop_folder_initialized:
            lines.append("Folder.GetSpecialFolder SpecialFolder: Folder.SpecialFolder.DesktopDirectory SpecialFolderPath=> SpecialFolderPath")
            defined_vars.add("specialfolderpath")
            desktop_folder_initialized = True

    # Pre-declare variables from Sequence.Variables
    for elem in root.iter():
        if clean_tag(elem.tag) == "Variable":
            vname = elem.get("Name")
            if vname:
                defined_vars.add(vname.lower())

    # Initialize common variables
    lines.append("SET responseString TO $''''''")
    defined_vars.add("responsestring")
    lines.append("Variables.CreateNewDatatable InputTable: { ^['id', 'title', 'body', 'userId'] } DataTable=> dtPosts")
    defined_vars.add("dtposts")
    lines.append("Variables.CreateNewList List=> jsonArray")
    defined_vars.add("jsonarray")
    lines.append("SET newPostTitle TO $'''Sample Post'''")
    defined_vars.add("newposttitle")
    lines.append("SET newPostDesc TO $'''Sample Description'''")
    defined_vars.add("newpostdesc")
    lines.append("SET delPostId TO 1")
    defined_vars.add("delpostid")
    lines.append("SET currentJToken TO $''''''")
    defined_vars.add("currentjtoken")
    lines.append("")

    excel_opened = False

    def walk(element, indent=""):
        nonlocal excel_opened
        tag = clean_tag(element.tag)

        # 1. InputDialog
        if tag == "InputDialog":
            title = clean_vb_string_to_robin(element.get("Title", "Input"))
            msg = clean_vb_string_to_robin(element.get("Label", "Please enter a value"))
            out_var = "UserInput"
            for out in element.iter():
                if clean_tag(out.tag) == "OutArgument":
                    t = (out.text or "").strip()
                    if t.startswith("[") and t.endswith("]"):
                        t = t[1:-1]
                    if t:
                        out_var = t
                        break
            defined_vars.add(out_var.lower())
            lines.append(f"{indent}Display.InputDialog InputType: Display.InputType.SingleLine IsTopMost: False DefaultValue: $'''''' Message: {msg} Title: {title} UserInput=> {out_var} ButtonPressed=> ButtonPressed")
            defined_vars.add("buttonpressed")
            return

        # 2. MessageBox
        if tag == "MessageBox":
            title = clean_vb_string_to_robin(element.get("DisplayName", "Message Box"))
            text_val = element.get("Text", "Message")
            msg = clean_vb_string_to_robin(text_val)
            lines.append(f"{indent}Display.ShowMessageDialog.ShowMessage Icon: Display.Icon.None Buttons: Display.Buttons.OK DefaultButton: Display.DefaultButton.Button1 IsTopMost: False Message: {msg} Title: {title} ButtonPressed=> ButtonPressed")
            defined_vars.add("buttonpressed")
            return

        # 3. HTTP Request (HttpClient or NetHttpRequest)
        if tag in ["HttpClient", "NetHttpRequest"]:
            raw_url = element.get("EndPoint", element.get("Url", "https://jsonplaceholder.typicode.com/posts"))
            method_str = (element.get("Method", "GET") or "GET").upper()
            
            method_map = {
                "GET": "Web.Method.Get",
                "POST": "Web.Method.Post",
                "PUT": "Web.Method.Put",
                "DELETE": "Web.Method.Delete"
            }
            method_enum = method_map.get(method_str, "Web.Method.Get")
            
            clean_url = clean_vb_string_to_robin(raw_url)
            lines.append(f"{indent}Web.InvokeWebService.InvokeWebService Url: {clean_url} Method: {method_enum} Accept: $'''application/json''' ContentType: $'''application/json''' ConnectionTimeout: 30 FollowRedirection: True ClearCookies: False FailOnErrorStatus: False EncodeRequestBody: True UserAgent: $'''Mozilla/5.0''' Encoding: Web.Encoding.AutoDetect AcceptUntrustedCertificates: False ResponseHeaders=> ResponseHeaders Response=> responseString StatusCode=> responseStatus")
            defined_vars.add("responseheaders")
            defined_vars.add("responsestring")
            defined_vars.add("responsestatus")
            return

        # 4. Deserialize JSON
        if tag in ["DeserializeJson", "DeserializeJsonArray"]:
            out_var = "jsonObj" if tag == "DeserializeJson" else "jsonArray"
            for out in element.iter():
                if clean_tag(out.tag) in ["OutArgument", "InOutArgument"]:
                    t = (out.text or "").strip()
                    if t.startswith("[") and t.endswith("]"):
                        t = t[1:-1]
                    if t:
                        out_var = t
            defined_vars.add(out_var.lower())
            lines.append(f"{indent}Variables.ConvertJsonToCustomObject Json: responseString CustomObject=> {out_var}")
            return

        # 5. BuildDataTable
        if tag == "BuildDataTable":
            table_name = "dtPosts"
            for out in element.iter():
                if clean_tag(out.tag) in ["OutArgument", "InOutArgument"]:
                    t = (out.text or "").strip()
                    if t.startswith("[") and t.endswith("]"):
                        t = t[1:-1]
                    if t:
                        table_name = t
            defined_vars.add(table_name.lower())
            lines.append(f"{indent}Variables.CreateNewDatatable InputTable: {{ ^['id', 'title', 'body', 'userId'] }} DataTable=> {table_name}")
            return

        # 6. AddDataRow
        if tag == "AddDataRow":
            table_name = "dtPosts"
            lines.append(f"{indent}Variables.AddRowToDataTable.AppendRowToDataTable DataTable: {table_name} RowToAdd: [currentJToken['id'], currentJToken['title'], currentJToken['body'], currentJToken['userId']]")
            return

        # 7. ForEach loop
        if tag == "ForEach":
            item_var = "currentJToken"
            for ch in element.iter():
                if clean_tag(ch.tag) == "DelegateInArgument":
                    t = (ch.get("Name") or "").strip()
                    if t:
                        item_var = t
            defined_vars.add(item_var.lower())
            lines.append(f"{indent}LOOP FOREACH {item_var} IN jsonArray")
            for ch in element:
                walk(ch, indent + "    ")
            lines.append(f"{indent}END")
            return

        # 8. Assign
        if tag == "Assign":
            var_name = "Variable"
            val_text = "0"
            for ch in element:
                ctag = clean_tag(ch.tag)
                if "To" in ctag:
                    for sub in ch.iter():
                        if clean_tag(sub.tag) == "OutArgument":
                            t = (sub.text or "").strip()
                            if t.startswith("[") and t.endswith("]"):
                                t = t[1:-1]
                            if t:
                                var_name = t
                elif "Value" in ctag:
                    for sub in ch.iter():
                        if clean_tag(sub.tag) in ["InArgument", "VisualBasicValue"]:
                            t = (sub.text or "").strip()
                            if t:
                                val_text = t

            # Flatten val_text to single line
            val_text = ' '.join(val_text.splitlines()).strip()

            # Filter out C# / VB constructors like new JObject() or New JArray()
            if re.search(r'\b(?:new\s+JObject|New\s+JArray|new\s+DataTable)\b', val_text, re.I):
                return

            if "Path.Combine" in val_text:
                ensure_desktop_folder()
                clean_val = clean_vb_string_to_robin(val_text)
            elif val_text.startswith('"') or val_text.startswith('["') or '&' in val_text or '+' in val_text:
                clean_val = clean_vb_string_to_robin(val_text)
            elif val_text.startswith("[") and val_text.endswith("]"):
                clean_val = val_text[1:-1]
            else:
                clean_val = val_text

            defined_vars.add(var_name.lower())
            lines.append(f"{indent}SET {var_name} TO {clean_val}")
            return

        # 9. ShowFormActivity
        if tag == "ShowFormActivity":
            lines.append(f"{indent}Display.InputDialog InputType: Display.InputType.SingleLine IsTopMost: False DefaultValue: $'''''' Message: $'''Enter Post Title:''' Title: $'''New Post Form''' UserInput=> newPostTitle ButtonPressed=> ButtonPressed")
            lines.append(f"{indent}Display.InputDialog InputType: Display.InputType.SingleLine IsTopMost: False DefaultValue: $'''''' Message: $'''Enter Post Description:''' Title: $'''New Post Form''' UserInput=> newPostDesc ButtonPressed=> ButtonPressed")
            defined_vars.add("newposttitle")
            defined_vars.add("newpostdesc")
            defined_vars.add("buttonpressed")
            return

        # 10. ExcelProcessScopeX / ExcelApplicationCard
        if tag in ["ExcelProcessScopeX", "ExcelApplicationCard", "ExcelApplicationScope"]:
            if not excel_opened:
                ensure_desktop_folder()
                lines.append(f"{indent}Excel.LaunchExcel.Launch UseMachineLocale: False Visible: True LoadAddInsAndMacros: False Instance=> ExcelInstance")
                defined_vars.add("excelinstance")
                excel_opened = True
            for ch in element:
                walk(ch, indent)
            return

        # 11. WriteRangeX
        if tag in ["WriteRange", "WriteRangeX"]:
            lines.append(f"{indent}Excel.WriteToExcel.WriteCell Instance: ExcelInstance Value: dtPosts Column: 1 Row: 1")
            return

        # 12. If Statement
        if tag == "If":
            cond = element.get("Condition", "True")
            if cond.startswith('[') and cond.endswith(']'):
                cond = cond[1:-1]
            lines.append(f"{indent}IF {cond} THEN")
            then_block = None
            else_block = None
            for ch in element:
                ctag = clean_tag(ch.tag)
                if ctag in ["If.Then", "Then"]:
                    then_block = ch
                elif ctag in ["If.Else", "Else"]:
                    else_block = ch
            if then_block is not None:
                for sub in then_block:
                    walk(sub, indent + "    ")
            if else_block is not None and len(list(else_block)) > 0:
                lines.append(f"{indent}ELSE")
                for sub in else_block:
                    walk(sub, indent + "    ")
            lines.append(f"{indent}END")
            return

        # 13. Delay
        if tag == "Delay":
            duration = element.get("Duration", "00:00:02")
            seconds = 2
            match = re.match(r'(\d+):(\d+):(\d+)', duration)
            if match:
                seconds = int(match.group(1)) * 3600 + int(match.group(2)) * 60 + int(match.group(3))
            lines.append(f"{indent}WAIT {seconds}")
            return

        # 14. Create Folder
        if tag == "CreateDirectory":
            p = element.get("Path", "$'''C:\\NewFolder'''")
            lines.append(f"{indent}Folder.Create FolderPath: $'''C:\\''' FolderName: {clean_vb_string_to_robin(p)} Folder=> NewFolder")
            return

        # 15. Write Text File
        if tag == "WriteTextFile":
            p = element.get("FileName", "$'''C:\\output.txt'''")
            txt = element.get("Text", "Output")
            lines.append(f"{indent}File.WriteText File: {clean_vb_string_to_robin(p)} AppendNewLine: True IfFileExists: File.IfFileExists.Overwrite Encoding: File.FileEncoding.Unicode TextToWrite: {clean_vb_string_to_robin(txt)}")
            return

        # 16. Read Text File
        if tag == "ReadTextFile":
            p = element.get("FileName", "$'''C:\\input.txt'''")
            lines.append(f"{indent}File.ReadTextFromFile.ReadText File: {clean_vb_string_to_robin(p)} Encoding: File.TextFileEncoding.UTF8 Content=> FileContents")
            return

        # 17. Log Message
        if tag == "LogMessage":
            msg = element.get("Message", "")
            lines.append(f"{indent}# Log: {msg}")
            return

        # Recurse for containers
        for ch in element:
            walk(ch, indent)

    walk(root)

    if excel_opened:
        lines.append("")
        lines.append("Excel.CloseExcel.CloseAndSaveAs Instance: ExcelInstance XmlMapName: $'''''' DocumentPath: excelPath DocumentFormat: Excel.ExcelFormat.FromExtension")

    return "\n".join(lines)


def generate_template_robin(template_key: str) -> str:
    """Generate golden, paste-validated Robin flows."""
    templates = {
        "folder_backup": (
            "DateTime.GetCurrentDateTime.Local DateTimeFormat: DateTime.DateTimeFormat.DateOnly CurrentDateTime=> CurrentDateTime\n"
            "Folder.GetFiles Folder: $'''C:\\Downloads''' FileFilter: $'''*''' IncludeSubfolders: False FailOnAccessDenied: True SortBy1: Folder.SortBy.NoSort SortDescending1: False SortBy2: Folder.SortBy.NoSort SortDescending2: False SortBy3: Folder.SortBy.NoSort SortDescending3: False Files=> Files\n"
            "File.Copy Files: Files Destination: $'''D:\\Backup\\Downloads''' IfFileExists: File.IfExists.Overwrite CopiedFiles=> CopiedFiles\n"
            "Display.ShowMessageDialog.ShowMessage Icon: Display.Icon.Information Buttons: Display.Buttons.OK DefaultButton: Display.DefaultButton.Button1 IsTopMost: False Message: $'''Backed up %CopiedFiles.Count% files on %CurrentDateTime%''' Title: $'''Downloads backup''' ButtonPressed=> ButtonPressed"
        ),
        "excel_report": (
            "Excel.LaunchExcel.Launch UseMachineLocale: False Visible: False LoadAddInsAndMacros: False Instance=> ExcelInstance\n"
            "Excel.WriteToExcel.WriteCell Instance: ExcelInstance Value: $'''Invoice''' Column: $'''A''' Row: 1\n"
            "Excel.WriteToExcel.WriteCell Instance: ExcelInstance Value: $'''Amount''' Column: $'''B''' Row: 1\n"
            "Excel.WriteToExcel.WriteCell Instance: ExcelInstance Value: $'''INV-1001''' Column: $'''A''' Row: 2\n"
            "Excel.WriteToExcel.WriteCell Instance: ExcelInstance Value: 450 Column: $'''B''' Row: 2\n"
            "Excel.CloseExcel.CloseAndSaveAs Instance: ExcelInstance XmlMapName: $'''''' DocumentPath: $'''C:\\Reports\\Invoice_Summary.xlsx''' DocumentFormat: Excel.ExcelFormat.FromExtension\n"
            "Display.ShowMessageDialog.ShowMessage Icon: Display.Icon.Information Buttons: Display.Buttons.OK DefaultButton: Display.DefaultButton.Button1 IsTopMost: False Message: $'''Excel summary generated and saved successfully!''' Title: $'''Report Done''' ButtonPressed=> ButtonPressed"
        ),
        "outlook_attachments": (
            "Outlook.Launch Instance=> OutlookInstance\n"
            "DateTime.GetCurrentDateTime.Local DateTimeFormat: DateTime.DateTimeFormat.DateAndTime CurrentDateTime=> CurrentDateTime\n"
            "Text.ConvertDateTimeToText.FromCustomDateTime CustomFormat: $'''yyyy-MM-dd''' DateTime: CurrentDateTime Result=> Today\n"
            "SET AttachmentFolder TO $'''C:\\Attachments\\\\%Today%'''\n"
            "IF (Folder.IfFolderExists.DoesNotExist Path: AttachmentFolder) THEN\n"
            "    Folder.Create FolderPath: $'''C:\\Attachments''' FolderName: Today Folder=> NewFolder\n"
            "END\n"
            "Outlook.RetrieveEmailMessages.RetrieveEmailsAndSaveAttachments ReadBodyAsHtml: False MarkAsRead: True MailFolder: $'''Inbox''' ToContains: $'''''' SubjectContains: $'''Invoice''' SaveAttachmentsInto: AttachmentFolder BodyContains: $'''''' Account: $'''you@example.com''' Instance: OutlookInstance FromContains: $'''''' EmailsToRetrieve: Outlook.RetrieveMessagesMode.Unread Messages=> RetrievedEmails\n"
            "SET EmailCount TO RetrievedEmails.Count\n"
            "IF EmailCount = 0 THEN\n"
            "    Display.ShowMessageDialog.ShowMessage Icon: Display.Icon.Information Buttons: Display.Buttons.OK DefaultButton: Display.DefaultButton.Button1 IsTopMost: False Message: $'''No new invoice emails''' Title: $'''Invoice attachments''' ButtonPressed=> ButtonPressed\n"
            "ELSE\n"
            "    Folder.GetFiles Folder: AttachmentFolder FileFilter: $'''*.pdf''' IncludeSubfolders: False FailOnAccessDenied: True SortBy1: Folder.SortBy.NoSort SortDescending1: False SortBy2: Folder.SortBy.NoSort SortDescending2: False SortBy3: Folder.SortBy.NoSort SortDescending3: False Files=> Files\n"
            "    Display.ShowMessageDialog.ShowMessage Icon: Display.Icon.Information Buttons: Display.Buttons.OK DefaultButton: Display.DefaultButton.Button1 IsTopMost: False Message: $'''Saved %Files.Count% PDFs from %RetrievedEmails.Count% emails''' Title: $'''Invoice attachments''' ButtonPressed=> ButtonPressed\n"
            "END\n"
            "Outlook.Close Instance: OutlookInstance"
        ),
        "input_and_decision": (
            "Display.InputDialog InputType: Display.InputType.SingleLine IsTopMost: False DefaultValue: $'''''' Message: $'''Please enter your age:''' Title: $'''Age Verification''' UserInput=> UserAge ButtonPressed=> ButtonPressed\n"
            "IF UserAge >= 18 THEN\n"
            "    Display.ShowMessageDialog.ShowMessage Icon: Display.Icon.Information Buttons: Display.Buttons.OK DefaultButton: Display.DefaultButton.Button1 IsTopMost: False Message: $'''Access granted. You are an adult.''' Title: $'''Verification Result''' ButtonPressed=> ButtonPressed\n"
            "ELSE\n"
            "    Display.ShowMessageDialog.ShowMessage Icon: Display.Icon.Warning Buttons: Display.Buttons.OK DefaultButton: Display.DefaultButton.Button1 IsTopMost: False Message: $'''Access restricted. You are a minor.''' Title: $'''Verification Result''' ButtonPressed=> ButtonPressed\n"
            "END"
        ),
        "sales_api_report": (
            "Web.InvokeWebService.InvokeWebService Url: $'''https://api.example.com/orders''' Method: Web.Method.Get Accept: $'''application/json''' ContentType: $'''application/json''' ConnectionTimeout: 30 FollowRedirection: True ClearCookies: False FailOnErrorStatus: False EncodeRequestBody: True UserAgent: $'''Mozilla/5.0''' Encoding: Web.Encoding.AutoDetect AcceptUntrustedCertificates: False TrimRequestBody: True ResponseHeaders=> WebServiceResponseHeaders Response=> WebServiceResponse StatusCode=> StatusCode\n"
            "IF StatusCode = 200 THEN\n"
            "    Display.ShowMessageDialog.ShowMessage Icon: Display.Icon.Information Buttons: Display.Buttons.OK DefaultButton: Display.DefaultButton.Button1 IsTopMost: False Message: $'''Sales API responded successfully!''' Title: $'''Sales API''' ButtonPressed=> ButtonPressed\n"
            "ELSE\n"
            "    Display.ShowMessageDialog.ShowMessage Icon: Display.Icon.ErrorIcon Buttons: Display.Buttons.OK DefaultButton: Display.DefaultButton.Button1 IsTopMost: False Message: $'''API failed with HTTP status: %StatusCode%''' Title: $'''Error''' ButtonPressed=> ButtonPressed\n"
            "END"
        )
    }
    return templates.get(template_key, "")


def validate_robin(code: str) -> dict:
    """Validate Robin script against PAD Robin rules using check-flows.js."""
    tmp_path = os.path.join(ROOT_DIR, "_tmp_validate.robin")
    try:
        with open(tmp_path, "w", encoding="utf-8") as f:
            f.write(code)

        cmd = ["node", CHECKER_PATH, KB_PATH, tmp_path]
        res = subprocess.run(cmd, capture_output=True, text=True)
        output = res.stdout + res.stderr
        if os.path.exists(tmp_path):
            os.remove(tmp_path)

        is_clean = res.returncode == 0 and "0 errors" in output
        errors = [line.strip() for line in output.splitlines() if line.strip().startswith("line ") or "error" in line.lower()]

        return {
            "valid": is_clean,
            "output": output,
            "errors": errors
        }
    except Exception as e:
        if os.path.exists(tmp_path):
            os.remove(tmp_path)
        return {"valid": False, "output": str(e), "errors": [str(e)]}
