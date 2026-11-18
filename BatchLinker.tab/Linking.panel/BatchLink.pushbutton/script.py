# -*- coding: utf-8 -*-
"""
Batch Model Linker for Revit (Excel Matrix Driven)
Automates linking N models based on an Excel spreadsheet with 2 tabs:
  Tab 1: 'Links'  - Model definitions, file paths, positioning, link types, and worksets
  Tab 2: 'Matrix' - Grid defining which models get linked into which host models

Key Capabilities:
1. Pure pyRevit (Zero external dependencies, zero extra add-ins required)
2. Works directly from Revit blank Home Screen ('zero-doc' context)
3. In-memory RAM optimization (Auto-unload links & CloseAllWorksets)
4. UI DialogBox & TaskDialog suppression
5. Transaction Warning suppression (IFailuresPreprocessor)
6. Full audit logging to file
7. Workset creation & assignment to both LinkType and LinkInstance
8. Sync & Relinquish support for central models
9. ACC / BIM360 cloud path support (BIM360:// and ACC:// paths accepted in Excel)
"""

__title__ = "Batch Linker"
__author__ = "Antigravity & BWT"
__context__ = "zero-doc"

import os
import gc
import sys
import datetime
from pyrevit import forms, script

import clr
clr.AddReference("System")
clr.AddReference("System.Xml")
clr.AddReference("System.IO.Compression.FileSystem")
clr.AddReference("RevitAPI")
clr.AddReference("RevitAPIUI")

from System import GC
import System.Diagnostics
from System.IO.Compression import ZipFile
from System.Xml import XmlDocument

from Autodesk.Revit.UI import TaskDialogResult
from Autodesk.Revit.UI.Events import DialogBoxShowingEventArgs, TaskDialogShowingEventArgs
from Autodesk.Revit.DB import (
    ModelPathUtils,
    OpenOptions,
    WorksetConfiguration,
    WorksetConfigurationOption,
    Transaction,
    FailureHandlingOptions,
    IFailuresPreprocessor,
    FailureProcessingResult,
    RevitLinkType,
    RevitLinkInstance,
    RevitLinkOptions,
    ImportPlacement,
    AttachmentType,
    Workset,
    WorksetKind,
    FilteredWorksetCollector,
    BuiltInParameter,
    SynchronizeWithCentralOptions,
    TransactWithCentralOptions,
    RelinquishOptions
)

# ----------------------------------------------------------------------
# 1. PURE .NET OPENXML EXCEL PARSER (NO THIRD PARTY PACKAGES NEEDED)
# ----------------------------------------------------------------------
class ExcelMatrixReader(object):
    """
    Parses an .xlsx file using standard .NET System.IO.Compression and System.Xml.
    Reads Tab 1 ('Links') and Tab 2 ('Matrix').
    """
    def __init__(self, excel_path):
        self.excel_path = excel_path
        self.models = {}         # key: Model_Code (e.g. 'M01') -> model dict
        self.matrix_rules = {}   # key: host_code -> list of link_codes to link
        self.shared_strings = []
        self._load()

    def _load(self):
        zip_file = ZipFile.OpenRead(self.excel_path)
        try:
            # 1. Read shared strings
            ss_entry = zip_file.GetEntry("xl/sharedStrings.xml")
            if ss_entry:
                s = ss_entry.Open()
                try:
                    xdoc = XmlDocument()
                    xdoc.Load(s)
                    for si in xdoc.GetElementsByTagName("si"):
                        self.shared_strings.append(si.InnerText)
                finally:
                    s.Close()

            # 2. Map sheet names from workbook.xml
            wb_entry = zip_file.GetEntry("xl/workbook.xml")
            wb_rels_entry = zip_file.GetEntry("xl/_rels/workbook.xml.rels")
            sheet_paths = {}

            if wb_entry and wb_rels_entry:
                rid_to_target = {}
                s = wb_rels_entry.Open()
                try:
                    rdoc = XmlDocument()
                    rdoc.Load(s)
                    for rel in rdoc.GetElementsByTagName("Relationship"):
                        rid_to_target[rel.GetAttribute("Id")] = rel.GetAttribute("Target")
                finally:
                    s.Close()

                s = wb_entry.Open()
                try:
                    wdoc = XmlDocument()
                    wdoc.Load(s)
                    for sheet in wdoc.GetElementsByTagName("sheet"):
                        name = sheet.GetAttribute("name").strip().lower()
                        rid = sheet.GetAttribute("r:id")
                        target = rid_to_target.get(rid, "")
                        if not target.startswith("xl/"):
                            target = "xl/" + target.lstrip("/")
                        sheet_paths[name] = target
                finally:
                    s.Close()

            # Identify sheets
            links_sheet_path = sheet_paths.get("links")
            matrix_sheet_path = sheet_paths.get("matrix")

            if not links_sheet_path:
                for k in sheet_paths:
                    if "link" in k:
                        links_sheet_path = sheet_paths[k]
                        break

            if not matrix_sheet_path:
                for k in sheet_paths:
                    if "matrix" in k:
                        matrix_sheet_path = sheet_paths[k]
                        break

            if not links_sheet_path or not matrix_sheet_path:
                raise Exception(
                    "The Excel file must contain two tabs named 'Links' and 'Matrix'. Found: {}".format(
                        ", ".join(sheet_paths.keys())
                    )
                )

            # 3. Parse 'Links' sheet
            links_rows = self._read_sheet_rows(zip_file, links_sheet_path)
            self._parse_links(links_rows)

            # 4. Parse 'Matrix' sheet
            matrix_rows = self._read_sheet_rows(zip_file, matrix_sheet_path)
            self._parse_matrix(matrix_rows)

        finally:
            zip_file.Dispose()

    def _read_sheet_rows(self, zip_file, sheet_path):
        entry = zip_file.GetEntry(sheet_path)
        if not entry:
            return []
        s = entry.Open()
        rows = []
        try:
            xdoc = XmlDocument()
            xdoc.Load(s)
            for r_node in xdoc.GetElementsByTagName("row"):
                cells = {}
                for c_node in r_node.GetElementsByTagName("c"):
                    r_ref = c_node.GetAttribute("r") # e.g. A1, B2
                    col = "".join([ch for ch in r_ref if ch.isalpha()]).upper()
                    c_type = c_node.GetAttribute("t")
                    val = ""
                    v_nodes = c_node.GetElementsByTagName("v")
                    if v_nodes.Count > 0:
                        raw = v_nodes[0].InnerText
                        if c_type == "s":
                            idx = int(raw)
                            val = self.shared_strings[idx] if idx < len(self.shared_strings) else ""
                        else:
                            val = raw
                    elif c_type == "inlineStr":
                        t_nodes = c_node.GetElementsByTagName("t")
                        if t_nodes.Count > 0:
                            val = t_nodes[0].InnerText
                    cells[col] = val.strip()
                rows.append(cells)
        finally:
            s.Close()
        return rows

    def _parse_links(self, rows):
        """
        Expects columns:
        A: Model Name, B: File Path, C: Positioning, D: Link Type, E: Workset
        """
        if not rows:
            return
        # Row 0 is header
        for row in rows[1:]:
            name = row.get("A", "").strip()
            if not name:
                continue
            path = row.get("B", "").strip()
            pos = row.get("C", "Origin").strip()
            l_type = row.get("D", "Overlay").strip()
            workset = row.get("E", "").strip()
            if not workset:
                workset = "Link_" + name

            self.models[name] = {
                "name": name,
                "path": path,
                "positioning": pos,
                "link_type": l_type,
                "workset": workset
            }

    def _parse_matrix(self, rows):
        """
        Row 0: Header with Col A: Model Name, Col B..: Model Names to link
        Rows 1+: Host row with Model Name in Col A and values (YES, Y, X, 1, etc.)
        """
        if len(rows) < 2:
            return
        header = rows[0]
        # Map column letters to target Model Names
        col_to_model = {}
        for col_letter, val in header.items():
            if col_letter == "A":
                continue
            col_to_model[col_letter] = val.strip()

        # Process each host row
        truthy_values = set(["yes", "y", "x", "1", "true"])
        for row in rows[1:]:
            host_name = row.get("A", "").strip()
            if not host_name:
                continue
            self.matrix_rules[host_name] = []

            for col_letter, target_name in col_to_model.items():
                cell_val = row.get(col_letter, "").strip().lower()
                if cell_val in truthy_values:
                    self.matrix_rules[host_name].append(target_name)

# ----------------------------------------------------------------------
# 2. CLOUD PATH HELPERS
# ----------------------------------------------------------------------
def is_cloud_path(path):
    """Returns True for BIM360:// or ACC:// cloud model paths."""
    lower = path.lower().strip()
    return lower.startswith("bim360://") or lower.startswith("acc://")

def cloud_safe_basename(path):
    """Returns the filename from both local and cloud paths."""
    # Normalise separators so os.path.basename works on cloud URIs too
    return os.path.basename(path.replace("\\", "/").rstrip("/"))

def cloud_safe_exists(path):
    """Returns True for cloud paths (existence verified by Revit on open) or local paths that exist."""
    if is_cloud_path(path):
        return True
    return os.path.exists(path)

# ----------------------------------------------------------------------
# 3. MEMORY MANAGEMENT
# ----------------------------------------------------------------------
class MemoryManager(object):
    @staticmethod
    def trim_memory():
        """Flushes Python GC, .NET CLR GC, and trims Revit's Windows Working Set."""
        try:
            gc.collect()
            GC.Collect()
            GC.WaitForPendingFinalizers()
            GC.Collect()
            
            proc = System.Diagnostics.Process.GetCurrentProcess()
            proc.MinWorkingSet = proc.MinWorkingSet
            proc.MaxWorkingSet = proc.MaxWorkingSet
        except Exception:
            pass

# ----------------------------------------------------------------------
# 4. AUDIT LOGGING
# ----------------------------------------------------------------------
class AuditLogger(object):
    def __init__(self, log_filepath):
        self.log_filepath = log_filepath
        self.current_host_model = "None"
        with open(self.log_filepath, "a") as f:
            f.write("\n" + "=" * 90 + "\n")
            f.write("BATCH LINKING AUDIT SESSION: {}\n".format(
                datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            ))
            f.write("=" * 90 + "\n")

    def log(self, event_type, details, action_taken):
        timestamp = datetime.datetime.now().strftime("%H:%M:%S")
        entry = "[{}] [{}] [Host: {}]\n  Details: {}\n  Action:  {}\n".format(
            timestamp, event_type, self.current_host_model, details, action_taken
        )
        with open(self.log_filepath, "a") as f:
            f.write(entry + "\n")

# ----------------------------------------------------------------------
# 5. WARNING & DIALOG SUPPRESSION
# ----------------------------------------------------------------------
class LoggingWarningSilencer(IFailuresPreprocessor):
    """Suppresses and logs all transaction warnings."""
    def __init__(self, logger):
        self.logger = logger

    def PreprocessFailures(self, failuresAccessor):
        for failure in failuresAccessor.GetFailureMessages():
            try:
                desc = failure.GetDescriptionText()
                elem_ids = [str(e_id.IntegerValue) for e_id in failure.GetFailingElementIds()]
                elem_str = " (Elements: {})".format(", ".join(elem_ids)) if elem_ids else ""
                self.logger.log(
                    "TRANSACTION WARNING",
                    desc + elem_str,
                    "Deleted / Suppressed"
                )
                failuresAccessor.DeleteWarning(failure)
            except Exception:
                pass
        return FailureProcessingResult.Continue


class LoggingDialogHandler(object):
    """Intercepts and auto-dismisses modal popups and TaskDialogs."""
    def __init__(self, logger):
        self.logger = logger

    def on_dialog_showing(self, sender, args):
        try:
            if isinstance(args, TaskDialogShowingEventArgs):
                dlg_id = args.DialogId or "Unknown TaskDialog"
                msg = args.Message.replace("\n", " ") if args.Message else "No message text"
                self.logger.log(
                    "UI TASK DIALOG",
                    "[ID: {}] {}".format(dlg_id, msg),
                    "Overridden with TaskDialogResult.Ok"
                )
                args.OverrideResult(TaskDialogResult.Ok)

            elif isinstance(args, DialogBoxShowingEventArgs):
                dlg_id = args.DialogId or "Windows Modal Dialog"
                self.logger.log(
                    "MODAL WINDOW",
                    "[ID: {}]".format(dlg_id),
                    "Overridden with IDOK (1)"
                )
                args.OverrideResult(1)
        except Exception:
            pass

# ----------------------------------------------------------------------
# 6. WORKSET & LINK HELPERS
# ----------------------------------------------------------------------
def get_or_create_workset(doc, workset_name):
    """Finds existing user workset or creates a new one (must be called inside an active transaction)."""
    user_worksets = FilteredWorksetCollector(doc).OfKind(WorksetKind.UserWorkset)
    for ws in user_worksets:
        if ws.Name.strip().lower() == workset_name.strip().lower():
            return ws.Id
    new_ws = Workset.Create(doc, workset_name)
    return new_ws.Id


def assign_workset_to_element(element, workset_id):
    """Sets ELEM_PARTITION_PARAM on an element."""
    try:
        param = element.get_Parameter(BuiltInParameter.ELEM_PARTITION_PARAM)
        if param and not param.IsReadOnly:
            param.Set(workset_id.IntegerValue)
    except Exception:
        pass


def get_existing_link_names(doc):
    """Returns a set of lowercase link type names already present in document."""
    from Autodesk.Revit.DB import FilteredElementCollector
    link_types = FilteredElementCollector(doc).OfClass(RevitLinkType).ToElements()
    return set(lt.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM).AsString().lower() 
               for lt in link_types if lt.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM))

# ----------------------------------------------------------------------
# 7. MAIN CONTROLLER
# ----------------------------------------------------------------------
def main():
    output = script.get_output()
    app = __revit__.Application

    # 1. Pick Excel Matrix File
    excel_path = forms.pick_file(
        file_ext="xlsx",
        title="Select Batch Link Matrix Excel File (.xlsx)"
    )
    if not excel_path:
        forms.alert("No Excel file selected. Operation cancelled.", exitscript=True)

    # 2. Parse Excel File
    try:
        matrix_data = ExcelMatrixReader(excel_path)
    except Exception as ex:
        forms.alert("Error reading Excel file:\n{}".format(str(ex)), exitscript=True)

    if not matrix_data.models:
        forms.alert("No models found in the 'Links' tab of the Excel file.", exitscript=True)

    # 3. Present Host Models to Process (Allows chunking)
    host_names = [name for name in matrix_data.matrix_rules.keys() if name in matrix_data.models]
    if not host_names:
        forms.alert("No host models found in the 'Matrix' tab that match 'Model Name' entries in the 'Links' tab.", exitscript=True)

    display_options = []
    for name in host_names:
        link_targets = matrix_data.matrix_rules.get(name, [])
        display_options.append("{} (Links to add: {})".format(name, len(link_targets)))

    selected_displays = forms.SelectFromList.show(
        display_options,
        title="Select Host Models to Process (Runs 100% in Background)",
        multiselect=True,
        button_name="Run Batch Linking"
    )
    if not selected_displays:
        forms.alert("No host models selected. Operation cancelled.", exitscript=True)

    selected_host_names = [d.split(" (Links to add:")[0].strip() for d in selected_displays]

    # 4. Setup Logging
    excel_dir = os.path.dirname(excel_path)
    log_dir = os.path.join(excel_dir, "Audit_Logs")
    if not os.path.exists(log_dir):
        os.makedirs(log_dir)
    log_filepath = os.path.join(log_dir, "batch_link_audit_{}.log".format(
        datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    ))
    logger = AuditLogger(log_filepath)

    # 5. Attach Global UI Event Handler
    dialog_handler = LoggingDialogHandler(logger)
    __revit__.DialogBoxShowing += dialog_handler.on_dialog_showing
    silencer = LoggingWarningSilencer(logger)

    print("=" * 80)
    print("STARTING EXCEL-DRIVEN BATCH LINKING PROCESS")
    print("Excel Configuration: {}".format(excel_path))
    print("Total Defined Models: {}".format(len(matrix_data.models)))
    print("Hosts to Process:     {}".format(len(selected_host_names)))
    print("Audit Log Path:       {}".format(log_filepath))
    print("=" * 80)

    total_hosts = len(selected_host_names)
    try:
        for idx, host_name in enumerate(selected_host_names, start=1):
            host_info = matrix_data.models[host_name]
            host_path = host_info["path"]
            logger.current_host_model = host_name
            print("\n[{}/{}] Processing Host: {}".format(idx, total_hosts, host_name))

            if not cloud_safe_exists(host_path):
                print("  [ERROR] Host file not found at: {}".format(host_path))
                logger.log("FILE NOT FOUND", "Host path does not exist: {}".format(host_path), "Skipped")
                continue

            # Configure Open Options (Open with All Worksets Closed for minimal RAM)
            open_opts = OpenOptions()
            ws_config = WorksetConfiguration(WorksetConfigurationOption.CloseAllWorksets)
            open_opts.SetOpenWorksetsConfiguration(ws_config)
            open_opts.Audit = False

            model_path_obj = ModelPathUtils.ConvertUserVisiblePathToModelPath(host_path)
            host_doc = None

            try:
                host_doc = app.OpenDocumentFile(model_path_obj, open_opts)
                if not host_doc:
                    print("  [ERROR] Could not open host model: {}".format(host_path))
                    logger.log("OPEN ERROR", "Could not open document", "Skipped")
                    continue

                is_workshared = host_doc.IsWorkshared
                existing_links = get_existing_link_names(host_doc)
                links_to_add = matrix_data.matrix_rules.get(host_name, [])

                linked_count = 0
                for target_name in links_to_add:
                    if target_name.lower() == host_name.lower():
                        continue # Skip self-linking

                    target_info = matrix_data.models.get(target_name)
                    if not target_info:
                        print("  [WARNING] Target model '{}' not found in Links tab.".format(target_name))
                        continue

                    link_path = target_info["path"]
                    link_name = target_info["name"]
                    link_filename = cloud_safe_basename(link_path)

                    if not cloud_safe_exists(link_path):
                        print("  [WARNING] Link file not found: {}".format(link_path))
                        logger.log("LINK FILE NOT FOUND", link_path, "Skipped")
                        continue

                    # Check if already linked
                    link_base = os.path.splitext(link_filename)[0].lower()
                    if link_base in existing_links or link_filename.lower() in existing_links:
                        continue

                    # Positioning & Attachment Type
                    pos_str = target_info["positioning"].strip().lower()
                    if "shared" in pos_str:
                        import_placement = ImportPlacement.Shared
                    elif "center" in pos_str:
                        import_placement = ImportPlacement.Centered
                    else:
                        import_placement = ImportPlacement.Origin # Default: Internal Origin

                    is_overlay = "attach" not in target_info["link_type"].strip().lower()
                    l_opts = RevitLinkOptions(not is_overlay) # False = Overlay, True = Attachment

                    l_type_element = None
                    try:
                        # 1. Create Link Type, Link Instance, and assign Worksets INSIDE a transaction
                        with Transaction(host_doc, "Batch Link " + link_name) as t_link:
                            t_link.Start()
                            fail_opts = t_link.GetFailureHandlingOptions()
                            fail_opts.SetFailuresPreprocessor(silencer)
                            t_link.SetFailureHandlingOptions(fail_opts)

                            # Workset
                            target_ws_name = target_info["workset"]
                            ws_id = get_or_create_workset(host_doc, target_ws_name) if is_workshared else None

                            # Create Link Type
                            l_model_path = ModelPathUtils.ConvertUserVisiblePathToModelPath(link_path)
                            link_type_result = RevitLinkType.Create(host_doc, l_model_path, l_opts)

                            if link_type_result and link_type_result.ElementId:
                                l_type = host_doc.GetElement(link_type_result.ElementId)
                                l_type_element = l_type

                                if hasattr(l_type, "AttachmentType"):
                                    l_type.AttachmentType = AttachmentType.Overlay if is_overlay else AttachmentType.Attachment

                                # Create Link Instance
                                l_inst = RevitLinkInstance.Create(host_doc, l_type.Id, import_placement)

                                # Assign Workset
                                if is_workshared and ws_id:
                                    assign_workset_to_element(l_type, ws_id)
                                    assign_workset_to_element(l_inst, ws_id)

                            t_link.Commit()

                        # 2. Outside all transactions: Unload link geometry from RAM!
                        if l_type_element:
                            try:
                                l_type_element.Unload(None)
                            except Exception:
                                pass
                            linked_count += 1

                    except Exception as link_ex:
                        logger.log(
                            "LINK CREATION ERROR",
                            "Failed linking {} into {}: {}".format(link_filename, host_name, str(link_ex)),
                            "Skipped"
                        )

                print("  Linked & Unloaded {} models onto respective worksets.".format(linked_count))

                # Save / Sync
                if is_workshared:
                    transact_opts = TransactWithCentralOptions()
                    sync_opts = SynchronizeWithCentralOptions()
                    relinquish_opts = RelinquishOptions(True)
                    sync_opts.SetRelinquishOptions(relinquish_opts)
                    sync_opts.Comment = "Antigravity Batch Link: Attached models from Excel matrix."
                    host_doc.SynchronizeWithCentral(transact_opts, sync_opts)
                    print("  Synchronized and relinquished all worksets successfully.")
                else:
                    host_doc.Save()
                    print("  Saved non-workshared document.")

            except Exception as doc_ex:
                print("  [ERROR] Exception processing {}: {}".format(host_name, str(doc_ex)))
                logger.log("DOCUMENT ERROR", str(doc_ex), "Aborted document")

            finally:
                if host_doc:
                    try:
                        host_doc.Close(False)
                    except Exception:
                        pass
                MemoryManager.trim_memory()

        print("\n" + "=" * 80)
        print("EXCEL BATCH LINKING COMPLETE!")
        print("Audit Log generated at:")
        print(log_filepath)
        print("=" * 80)
        forms.alert("Batch linking complete!\n\nAudit log saved to:\n{}".format(log_filepath), title="Success")

    finally:
        __revit__.DialogBoxShowing -= dialog_handler.on_dialog_showing


if __name__ == "__main__":
    main()
