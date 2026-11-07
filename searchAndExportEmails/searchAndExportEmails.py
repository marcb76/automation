import argparse
import datetime
import os
import re
import shutil
import sys
import time
import win32com.client

# ==============================================================================
# DEFAULT CONFIGURATION (Fallbacks)
# ==============================================================================
DEFAULT_PST_FOLDER = "Outlook Data File"
DEFAULT_FILTER_EMAILS = ""
DEFAULT_FILTER_MODE = "both"  # Options: inbound, outbound, both
DEFAULT_OUTPUT_DIR = r"C:\Users\marcb\OneDrive\Personal Managed\DevEnv\automation\searchAndExportEmails\exports"


def print_header():
    """Prints the script header with metadata, description, and objective."""
    print("==================================================")
    print(" Script Name : searchAndExportEmails.py")
    print(" Version     : 1.0.29")
    print(" Created     : November 2026")
    print(" Author      : Marc Bonet")
    print(" Description : Automated search and export utility for Outlook desktop PST files.")
    print(" Objective   : Robust output directory clearing with interactive user prompt")
    print("               on file-locking (Access Denied) errors.")
    print("==================================================\n")


def sanitize_filename(filename):
    """Cleans email subjects, attachment names, or folder names for Windows paths."""
    if not filename:
        return "unnamed_file"
    sanitized = re.sub(r'[\\/*?:"<>|\r\n\t]', "_", filename).strip()
    sanitized = sanitized.rstrip(". ")
    return sanitized if sanitized else "unnamed_file"


def get_safe_attachment_path(att_dir, att_name):
    """Generates a safe path for Outlook attachment extraction within Windows path limits."""
    clean_name = sanitize_filename(att_name)
    abs_dir = os.path.abspath(att_dir)
    full_path = os.path.join(abs_dir, clean_name)

    if len(full_path) > 230:
        name_part, ext = os.path.splitext(clean_name)
        max_name_len = max(10, 230 - len(abs_dir) - len(ext) - 5)
        clean_name = name_part[:max_name_len] + ext
        full_path = os.path.join(abs_dir, clean_name)

    return full_path


def is_real_attachment(att):
    """Returns True only if the attachment is a true standalone file (excluding Outlook signature graphics)."""
    filename = getattr(att, "FileName", "") or ""
    if not filename:
        return False

    # 1. Exclude Outlook auto-generated signature/inline images (e.g. image001.png, image12345.jpg)
    if re.match(r'^image\d+.*\.(png|jpg|jpeg|gif|bmp)$', filename, re.IGNORECASE):
        return False

    try:
        # 2. Check ContentID: If present and not empty, it is an inline embedded resource
        content_id = getattr(att, "ContentID", None)
        if content_id and str(content_id).strip():
            return False
    except Exception:
        pass

    try:
        # 3. Check PR_RENDERING_POSITION (0x370B0003): -1 means standalone attachment.
        pa = att.PropertyAccessor
        rendering_pos = pa.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x370B0003")
        if rendering_pos != -1:
            return False
    except Exception:
        pass

    return True


def folder_should_process_items(folder_name, current_path, target_regex_str):
    """Determines if the folder's items should be processed based on folderRegex."""
    if not target_regex_str:
        return True

    patterns = target_regex_str.split("|")
    for pat in patterns:
        pat = pat.strip()
        if not pat:
            continue
        if re.search(pat, folder_name, re.IGNORECASE):
            return True
        if re.search(pat, current_path, re.IGNORECASE):
            return True
    return False


def remove_tree_with_retry(dir_path, retries=5, delay=0.5):
    """Robustly deletes a directory tree with retries to handle Windows file locking."""
    for attempt in range(retries):
        try:
            if os.path.exists(dir_path):
                shutil.rmtree(dir_path, ignore_errors=False)
            return True
        except Exception:
            if attempt < retries - 1:
                time.sleep(delay)
            else:
                try:
                    shutil.rmtree(dir_path, ignore_errors=True)
                except Exception:
                    pass
    return not os.path.exists(dir_path)


def process_folder(folder, filter_emails_list, base_output_dir, temp_work_dir, dry_run, 
                   filter_mode, subject_regex, start_date, end_date, folder_regex, 
                   current_rel_path, extract_attachments, convert_to_pdf, word_app, stats):
    """Recursively traverses Outlook folders, filters emails, and exports each email into
    its own dedicated chronological folder containing '_email.pdf' and its true attachments.
    """
    folder_name = getattr(folder, "Name", "Unknown")
    
    sanitized_folder_name = sanitize_filename(folder_name)
    next_rel_path = os.path.join(current_rel_path, sanitized_folder_name) if current_rel_path else sanitized_folder_name
    
    process_items = folder_should_process_items(folder_name, next_rel_path, folder_regex)

    if process_items:
        print(f"[*] Analyzing folder: '{folder_name}' -> Target path: '{next_rel_path}'")
        try:
            items = folder.Items
            for item in items:
                if hasattr(item, "Class") and item.Class == 43:
                    stats["processed"] += 1

                    # STEP 1: Date Range Filter
                    received_time = getattr(item, "ReceivedTime", None)
                    if start_date or end_date:
                        if not received_time:
                            continue
                        item_date = datetime.datetime(
                            received_time.year, received_time.month, received_time.day
                        ).date()
                        
                        if start_date and item_date < start_date:
                            continue
                        if end_date and item_date > end_date:
                            continue

                    # STEP 2: Subject Regex Filter
                    subject = getattr(item, "Subject", "No_Subject") or "No_Subject"
                    if subject_regex:
                        if not re.search(subject_regex, subject, re.IGNORECASE):
                            continue

                    # STEP 3: Target Email Addresses & Directional Filter
                    sender = getattr(item, "SenderEmailAddress", "") or ""
                    recipients_list = []
                    try:
                        if hasattr(item, "Recipients"):
                            for r in item.Recipients:
                                if hasattr(r, "Address") and r.Address:
                                    recipients_list.append(r.Address)
                    except Exception:
                        pass
                    
                    recipients_str = "; ".join(recipients_list)
                    sender_lower = sender.lower()
                    recipients_lower = recipients_str.lower()

                    match_sender = any(e in sender_lower for e in filter_emails_list)
                    match_recipient = any(e in recipients_lower for e in filter_emails_list)

                    match_found = False
                    if filter_mode == "inbound":
                        match_found = match_sender
                    elif filter_mode == "outbound":
                        match_found = match_recipient
                    elif filter_mode == "both":
                        match_found = match_sender or match_recipient

                    if not match_found:
                        continue

                    # ==========================================================
                    # ISOLATED TEMP PROCESSING & DEDICATED FOLDER EXPORT
                    # ==========================================================
                    stats["found"] += 1
                    stats["folder_counts"][next_rel_path] = stats["folder_counts"].get(next_rel_path, 0) + 1

                    date_str = received_time.strftime("%Y%m%d_%H%M%S") if received_time else "NO_DATE"
                    clean_subject = sanitize_filename(subject)[:40]

                    item_temp_dir = os.path.join(temp_work_dir, f"item_{stats['found']}")
                    os.makedirs(item_temp_dir, exist_ok=True)
                    
                    temp_msg_path = os.path.join(item_temp_dir, "message.msg")
                    temp_pdf_path = os.path.join(item_temp_dir, "_email.pdf")
                    temp_html_path = os.path.join(item_temp_dir, "message.html")
                    temp_att_dir = os.path.join(item_temp_dir, "attachments")

                    action_label = "Would export" if dry_run else "Exporting"
                    print(
                        f'  [{ "SIMULATION" if dry_run else "REAL" }] {action_label} -> '
                        f'Subject: "{subject[:35]}..." (Date: {date_str})'
                    )

                    if not dry_run:
                        try:
                            parent_folder_dir = os.path.join(base_output_dir, next_rel_path)
                            os.makedirs(parent_folder_dir, exist_ok=True)

                            # Dedicated chronological email folder name
                            email_folder_name = f"{date_str}_{clean_subject}"
                            email_folder_path = os.path.join(parent_folder_dir, email_folder_name)

                            counter = 1
                            while os.path.exists(email_folder_path):
                                email_folder_name = f"{date_str}_{clean_subject}_{counter}"
                                email_folder_path = os.path.join(parent_folder_dir, email_folder_name)
                                counter += 1

                            os.makedirs(email_folder_path, exist_ok=True)

                            # 1. Save temporary .msg
                            item.SaveAs(temp_msg_path, 3)

                            # 2. Extract True Attachments into temporary attachments directory
                            extracted_real_att_count = 0
                            if extract_attachments and hasattr(item, "Attachments") and item.Attachments.Count > 0:
                                real_attachments = [a for a in item.Attachments if is_real_attachment(a)]
                                if real_attachments:
                                    os.makedirs(temp_att_dir, exist_ok=True)
                                    for att in real_attachments:
                                        try:
                                            att_name = getattr(att, "FileName", "unnamed_attachment")
                                            att_path = get_safe_attachment_path(temp_att_dir, att_name)
                                            att.SaveAsFile(att_path)
                                            extracted_real_att_count += 1
                                        except Exception as ex:
                                            print(f"      [!] Error saving attachment '{getattr(att, 'FileName', '')}': {ex}")

                            # 3. Convert to PDF using Word via temporary HTML
                            if convert_to_pdf and word_app:
                                doc = None
                                try:
                                    item.SaveAs(temp_html_path, 5)
                                    doc = word_app.Documents.Open(temp_html_path, ReadOnly=True, Visible=False)
                                    doc.ExportAsFixedFormat(
                                        OutputFileName=temp_pdf_path,
                                        ExportFormat=17,
                                        OpenAfterExport=False,
                                        OptimizeFor=0,
                                        Range=0,
                                        Item=0
                                    )
                                    doc.Close(False)
                                    doc = None
                                except Exception as ex:
                                    print(f"      [!] Error converting to PDF: {ex}")
                                    if doc:
                                        try:
                                            doc.Close(False)
                                        except Exception:
                                            pass

                            # 4. Move _email.pdf into the dedicated email folder (Prefix '_' guarantees top position)
                            final_pdf_dest = os.path.join(email_folder_path, "_email.pdf")
                            if os.path.exists(temp_pdf_path):
                                shutil.move(temp_pdf_path, final_pdf_dest)

                            # 5. Move extracted true attachments directly into the dedicated email folder
                            if extracted_real_att_count > 0 and os.path.exists(temp_att_dir):
                                for file_name in os.listdir(temp_att_dir):
                                    src_file = os.path.join(temp_att_dir, file_name)
                                    dst_file = os.path.join(email_folder_path, file_name)
                                    if os.path.isfile(src_file):
                                        shutil.move(src_file, dst_file)
                                print(f"    [Attachments] Extracted {extracted_real_att_count} true attachment(s).")

                            # 6. Clean up individual item temporary folder
                            remove_tree_with_retry(item_temp_dir)
                            stats["exported"] += 1

                        except Exception as ex:
                            print(f"    [!] Error processing item: {ex}")
                            stats["errors"] += 1
                            if os.path.exists(item_temp_dir):
                                remove_tree_with_retry(item_temp_dir)

        except Exception as e:
            print(f"  [!] Error processing items in '{folder_name}': {e}")
    else:
        print(f"[-] Skipping folder (Ignored by filter): '{folder_name}' -> Target path: '{next_rel_path}'")

    try:
        for subfolder in folder.Folders:
            process_folder(
                subfolder, filter_emails_list, base_output_dir, temp_work_dir, dry_run, 
                filter_mode, subject_regex, start_date, end_date, folder_regex, 
                next_rel_path, extract_attachments, convert_to_pdf, word_app, stats
            )
    except Exception as e:
        print(f"  [!] Error accessing subfolders of '{folder_name}': {e}")


def run_export(pst_folder_name, filter_emails_arg, output_dir, dry_run, filter_mode_arg,
               subject_regex, start_date_arg, end_date_arg, folder_regex, clear_exports,
               extract_attachments, convert_to_pdf):
    
    start_dt = datetime.datetime.now()
    execution_date = start_dt.strftime("%Y-%m-%d")
    start_time_str = start_dt.strftime("%H:%M:%S")

    pst_folder_name = pst_folder_name or DEFAULT_PST_FOLDER
    output_dir = output_dir or DEFAULT_OUTPUT_DIR
    raw_emails = filter_emails_arg or DEFAULT_FILTER_EMAILS
    filter_mode = (filter_mode_arg or DEFAULT_FILTER_MODE).lower()

    if filter_mode not in ["inbound", "outbound", "both"]:
        print(f"[X] Invalid filter mode '{filter_mode}'. Valid options: inbound, outbound, both.")
        sys.exit(1)

    start_date = None
    if start_date_arg:
        try:
            start_date = datetime.datetime.strptime(start_date_arg, "%Y-%m-%d").date()
        except ValueError:
            print(f"[X] Invalid startDate format '{start_date_arg}'. Expected YYYY-MM-DD.")
            sys.exit(1)

    end_date = None
    if end_date_arg:
        try:
            end_date = datetime.datetime.strptime(end_date_arg, "%Y-%m-%d").date()
        except ValueError:
            print(f"[X] Invalid endDate format '{end_date_arg}'. Expected YYYY-MM-DD.")
            sys.exit(1)

    filter_emails_list = [e.strip().lower() for e in raw_emails.split(",") if e.strip()]
    mode_str = "SIMULATION (Dry Run)" if dry_run else "REAL EXECUTION"

    timestamp_str = start_dt.strftime("%Y%m%d_%H%M%S")
    temp_work_dir = os.path.join(output_dir, f"_temp_work_{timestamp_str}")

    print("--- STARTING OUTLOOK SEARCH & EXPORT ---")
    print(f"Execution Mode   : {mode_str}")
    print(f"PST Target Store : {pst_folder_name}")
    print(f"Filter Emails    : {filter_emails_list}")
    print(f"Filter Mode      : {filter_mode.upper()}")
    if subject_regex:
        print(f"Subject Regex    : {subject_regex}")
    if start_date:
        print(f"Start Date       : {start_date}")
    if end_date:
        print(f"End Date         : {end_date}")
    if folder_regex:
        print(f"Folder Regex     : {folder_regex}")
    print(f"Extract Attach.  : {extract_attachments}")
    print(f"Convert to PDF   : {convert_to_pdf}")
    print(f"Clear Exports    : {clear_exports}")
    print(f"Output Directory : {output_dir}")
    print(f"Temp Workspace   : {temp_work_dir}")
    print("----------------------------------------")

    if clear_exports:
        if dry_run:
            print(f"[*] [SIMULATION] Would clear output directory: {output_dir}")
        else:
            if os.path.exists(output_dir):
                print(f"[*] Clearing output directory: {output_dir}")
                cleared = False
                while not cleared:
                    if remove_tree_with_retry(output_dir, retries=3, delay=0.5):
                        cleared = True
                    else:
                        print(f"\n[!] Access Denied: Could not clear output directory (files may be locked or open).")
                        choice = input("    Please delete/unlock the output directory and press [E] to Export to continue, or [C] to Cancel: ").strip().lower()
                        if choice in ["c", "cancel", "abort"]:
                            print("[*] Aborted by user.")
                            sys.exit(0)
                        elif choice in ["e", "export", ""]:
                            print("[*] Retrying clearing output directory...")
                        else:
                            print("[*] Invalid choice. Retrying clearing output directory...")
                
                os.makedirs(output_dir, exist_ok=True)
            else:
                os.makedirs(output_dir, exist_ok=True)

    if not dry_run:
        os.makedirs(output_dir, exist_ok=True)
        os.makedirs(temp_work_dir, exist_ok=True)

    print("[*] Connecting to active Outlook instance (MAPI)...")
    try:
        outlook = win32com.client.Dispatch("Outlook.Application")
    except Exception as e:
        print(f"[X] Could not connect to Outlook. Ensure Outlook is running. Error: {e}")
        sys.exit(1)

    word_app = None
    if convert_to_pdf and not dry_run:
        try:
            word_app = win32com.client.Dispatch("Word.Application")
            word_app.Visible = False
            word_app.DisplayAlerts = 0
        except Exception as e:
            print(f"[!] Could not initialize Word for PDF conversion: {e}")

    namespace = outlook.GetNamespace("MAPI")
    target_pst = None
    available_stores = []
    for folder in namespace.Folders:
        available_stores.append(folder.Name)
        if pst_folder_name.lower() in folder.Name.lower():
            target_pst = folder
            break

    if not target_pst:
        print(f"[X] Could not find any store/PST matching '{pst_folder_name}'.")
        print("    Available stores/PSTs in your current Outlook session:")
        for store in available_stores:
            print(f"    - {store}")
        sys.exit(1)

    print(f"[✓] Store found: '{target_pst.Name}'\n")

    stats = {
        "processed": 0,
        "found": 0,
        "exported": 0,
        "errors": 0,
        "folder_counts": {}
    }
    
    try:
        process_folder(
            target_pst, filter_emails_list, output_dir, temp_work_dir, dry_run, 
            filter_mode, subject_regex, start_date, end_date, folder_regex, "", 
            extract_attachments, convert_to_pdf, word_app, stats
        )
    finally:
        # 1. Quit Word application first to release all file locks
        if word_app:
            try:
                word_app.Quit()
            except Exception:
                pass
            word_app = None
            time.sleep(0.5)

        # 2. Guarantee complete cleanup of the temporary workspace with retries
        if os.path.exists(temp_work_dir):
            print(f"[*] Cleaning up temporary workspace: {temp_work_dir}")
            remove_tree_with_retry(temp_work_dir)

    end_dt = datetime.datetime.now()
    end_time_str = end_dt.strftime("%H:%M:%S")
    elapsed_delta = end_dt - start_dt
    
    tot_seconds = int(elapsed_delta.total_seconds())
    hours, remainder = divmod(tot_seconds, 3600)
    minutes, seconds = divmod(remainder, 60)
    elapsed_str = f"{hours:02d}:{minutes:02d}:{seconds:02d}"

    print("\n--- EXECUTION SUMMARY ---")
    print(f"Execution Date              : {execution_date}")
    print(f"Start Time                  : {start_time_str}")
    print(f"End Time                    : {end_time_str}")
    print(f"Elapsed Time                : {elapsed_str}")
    print(f'Total emails processed      : {stats["processed"]}')
    print(f'Total matching emails found : {stats["found"]}')
    if not dry_run:
        print(f'Total emails saved          : {stats["exported"]}')
        print(f'Total errors encountered    : {stats["errors"]}')
    else:
        print("Mode was Dry Run: No files were written to disk.")
    
    print("\nFolder Breakdown (Matching Emails per Folder):")
    if stats["folder_counts"]:
        for folder_path, count in sorted(stats["folder_counts"].items()):
            print(f"  - {folder_path}: {count}")
    else:
        print("  (No matching emails found in any folder)")
    print("-----------------------------------------")

    summary_lines = [
        "==================================================",
        "          EXECUTION SUMMARY REPORT",
        " Script Name : searchAndExportEmails.py",
        " Version     : 1.0.29",
        "==================================================",
        "",
        "--- TIMESTAMPS & DURATION ---",
        f"Execution Date      : {execution_date}",
        f"Start Time          : {start_time_str}",
        f"End Time            : {end_time_str}",
        f"Elapsed Time        : {elapsed_str} ({tot_seconds} seconds)",
        "",
        "--- INITIAL PARAMETERS & VARIABLES ---",
        f"Execution Mode      : {mode_str}",
        f"PST Target Store    : {pst_folder_name}",
        f"Filter Emails       : {filter_emails_list}",
        f"Filter Mode         : {filter_mode.upper()}",
        f"Start Date          : {start_date or 'None'}",
        f"End Date            : {end_date or 'None'}",
        f"Subject Regex       : {subject_regex or 'None'}",
        f"Folder Regex        : {folder_regex or 'None'}",
        f"Extract Attach.     : {extract_attachments}",
        f"Convert to PDF      : {convert_to_pdf}",
        f"Clear Exports       : {clear_exports}",
        f"Output Directory    : {output_dir}",
        "",
        "--- FINAL RESULTS & FOLDER BREAKDOWN ---",
        f"Total Emails Processed  : {stats['processed']}",
        f"Total Emails Found      : {stats['found']}",
        f"Total Emails Exported   : {stats['exported'] if not dry_run else '0 (Dry Run)'}",
        f"Total Errors Encountered: {stats['errors'] if not dry_run else '0 (Dry Run)'}",
        "",
        "Folder Breakdown (Matching Emails per Folder):"
    ]

    if stats["folder_counts"]:
        for folder_path, count in sorted(stats["folder_counts"].items()):
            summary_lines.append(f"  - {folder_path}: {count}")
    else:
        summary_lines.append("  (No matching emails found in any folder)")

    summary_lines.append("==================================================")
    summary_content = "\n".join(summary_lines)

    summary_filename = "execution_summary.txt"
    summary_filepath = os.path.join(output_dir, summary_filename)

    if not dry_run:
        try:
            if not os.path.exists(output_dir):
                os.makedirs(output_dir, exist_ok=True)
            with open(summary_filepath, "w", encoding="utf-8") as sf:
                sf.write(summary_content)
            print(f"[✓] Execution summary report saved to: {summary_filepath}")
        except Exception as ex:
            print(f"[!] Error writing execution summary file: {ex}")
    else:
        print(f"[*] [SIMULATION] Would write execution summary report to: {summary_filepath}")


def main():
    print_header()

    parser = argparse.ArgumentParser(
        description="Automated search and export utility for Outlook desktop PST files.",
        add_help=True,
    )

    parser.add_argument("--dryRun", action="store_true", help="Simulation mode.")
    parser.add_argument("--clearExports", action="store_true", help="Clear output directory.")
    parser.add_argument("--extractAttachments", action="store_true", help="Extract attachments.")
    parser.add_argument("--convertToPDF", action="store_true", help="Convert to PDF.")
    parser.add_argument("--pstFolder", type=str, default=None, help="PST store name.")
    parser.add_argument("--filterEmails", type=str, default=None, help="Email filter list.")
    parser.add_argument("--filterMode", type=str, choices=["inbound", "outbound", "both"], default=None, help="Filter mode.")
    parser.add_argument("--subjectRegex", type=str, default=None, help="Subject regex.")
    parser.add_argument("--startDate", type=str, default=None, help="Start date YYYY-MM-DD.")
    parser.add_argument("--endDate", type=str, default=None, help="End date YYYY-MM-DD.")
    parser.add_argument("--folderRegex", type=str, default=None, help="Folder regex path filter.")
    parser.add_argument("--outputDir", type=str, default=None, help="Output directory path.")

    args = parser.parse_args()

    run_export(
        pst_folder_name=args.pstFolder,
        filter_emails_arg=args.filterEmails,
        output_dir=args.outputDir,
        dry_run=args.dryRun,
        filter_mode_arg=args.filterMode,
        subject_regex=args.subjectRegex,
        start_date_arg=args.startDate,
        end_date_arg=args.endDate,
        folder_regex=args.folderRegex,
        clear_exports=args.clearExports,
        extract_attachments=args.extractAttachments,
        convert_to_pdf=args.convertToPDF,
    )


if __name__ == "__main__":
    main()