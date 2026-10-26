#!/usr/bin/env python3
"""
Eloqua Form JSON Builder
------------------------
Tkinter application to generate a valid Eloqua Form create payload
(POST /api/rest/2.0/assets/form) from an Excel or CSV file of field definitions.

Features:
  - Load Excel (.xlsx/.xls) or CSV
  - Map columns to Eloqua FormField properties
  - Auto-generate negative temporary IDs required by Eloqua
  - Optional submit button + basic processing step
  - Live JSON preview with validation and pretty-print
  - Edit JSON freely, re-validate, and copy to clipboard
"""

import json
import os
import re
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
from typing import Any, Dict, List, Optional

try:
    import pandas as pd
except ImportError:
    pd = None

# ---------------------------------------------------------------------------
# Eloqua helpers
# ---------------------------------------------------------------------------

DEFAULT_STYLE = '{"fieldSize":"large","labelPosition":"top"}'
DEFAULT_FORM_STYLE = '{"fieldSize":"medium","labelPosition":"top"}'
SUBMIT_STYLE = '{"fieldSize":"large","submitButtonStyleType":"standard","submitButtonWidthStyle":"100px","submitButtonHeightStyle":"24px"}'

# Common column name aliases (case-insensitive)
COLUMN_ALIASES = {
    "name": ["name", "field name", "label", "field_label", "displayname", "display name"],
    "htmlName": ["htmlname", "html_name", "html name", "htmlname", "fieldname", "field_name", "api name", "apiname"],
    "dataType": ["datatype", "data_type", "data type", "type"],
    "displayType": ["displaytype", "display_type", "display type", "uitype", "ui type", "control"],
    "instructions": ["instructions", "helptext", "help text", "description", "hint"],
    "defaultValue": ["defaultvalue", "default_value", "default value", "default"],
    "required": ["required", "isrequired", "is_required", "mandatory"],
    "createdFromContactFieldId": ["createdfromcontactfieldid", "contactfieldid", "contact_field_id", "contact field id"],
    "fieldMergeId": ["fieldmergeid", "field_merge_id", "field merge id"],
}


def normalize_header(h: str) -> str:
    return re.sub(r"[\s_]+", "", str(h).strip().lower())


def resolve_column_map(headers: List[str]) -> Dict[str, str]:
    """Map canonical Eloqua keys -> actual column name in the file."""
    norm_to_orig = {normalize_header(h): h for h in headers}
    mapping: Dict[str, str] = {}
    for canon, aliases in COLUMN_ALIASES.items():
        for alias in aliases:
            key = normalize_header(alias)
            if key in norm_to_orig:
                mapping[canon] = norm_to_orig[key]
                break
    return mapping


def sanitize_html_name(value: str) -> str:
    """Produce a safe htmlName (letters, numbers, underscore)."""
    s = re.sub(r"[^a-zA-Z0-9_]", "", str(value).strip().replace(" ", "_"))
    if not s:
        s = "field"
    if s[0].isdigit():
        s = "f_" + s
    return s


def to_bool(val: Any) -> bool:
    if isinstance(val, bool):
        return val
    if val is None or (isinstance(val, float) and pd is not None and pd.isna(val)):
        return False
    return str(val).strip().lower() in {"1", "true", "yes", "y", "required", "mandatory"}


def build_validation_required(temp_id: int) -> Dict[str, Any]:
    return {
        "id": str(temp_id),
        "type": "FieldValidation",
        "message": "This field is required",
        "condition": {"type": "IsRequiredCondition"},
    }


def build_form_field(
    row: Dict[str, Any],
    col_map: Dict[str, str],
    next_id: int,
) -> tuple[Dict[str, Any], int]:
    """Build one FormField element and return (element, next available negative id)."""
    def get(canon: str, default: Any = "") -> Any:
        col = col_map.get(canon)
        if col is None:
            return default
        val = row.get(col, default)
        if pd is not None and isinstance(val, float) and pd.isna(val):
            return default
        return val if val is not None else default

    name = str(get("name") or "Untitled Field").strip()
    html_name = str(get("htmlName") or "").strip()
    if not html_name:
        html_name = sanitize_html_name(name)
    else:
        html_name = sanitize_html_name(html_name)

    data_type = str(get("dataType") or "text").strip().lower() or "text"
    display_type = str(get("displayType") or "text").strip().lower() or "text"
    # Eloqua common displayType values: text, textArea, singleSelect, multiSelect, checkbox, radio, hidden, submit, ...
    if display_type in ("textbox", "input"):
        display_type = "text"
    if display_type in ("textarea",):
        display_type = "textArea"

    element: Dict[str, Any] = {
        "type": "FormField",
        "id": str(next_id),
        "name": name,
        "htmlName": html_name,
        "dataType": data_type,
        "displayType": display_type,
        "style": DEFAULT_STYLE,
        "validations": [],
    }
    next_id -= 1

    instructions = str(get("instructions") or "").strip()
    if instructions:
        element["instructions"] = instructions

    default_value = get("defaultValue")
    if default_value not in ("", None):
        element["defaultValue"] = str(default_value)

    contact_field = str(get("createdFromContactFieldId") or "").strip()
    if contact_field:
        element["createdFromContactFieldId"] = contact_field

    merge_id = str(get("fieldMergeId") or "").strip()
    if merge_id:
        element["fieldMergeId"] = merge_id

    if to_bool(get("required")):
        element["validations"].append(build_validation_required(next_id))
        next_id -= 1

    return element, next_id


def build_submit_button(next_id: int) -> tuple[Dict[str, Any], int]:
    element = {
        "type": "FormField",
        "id": str(next_id),
        "name": "Submit",
        "htmlName": "submit",
        "dataType": "text",
        "displayType": "submit",
        "style": SUBMIT_STYLE,
        "validations": [],
    }
    return element, next_id - 1


def build_basic_create_update_step(
    email_field_id: Optional[str],
    next_id: int,
) -> tuple[Dict[str, Any], int]:
    """Optional FormStepCreateUpdateContactFromFormField (common processing step)."""
    step: Dict[str, Any] = {
        "type": "FormStepCreateUpdateContactFromFormField",
        "id": str(next_id),
        "name": "Create / Update Contact",
        "execute": "always",
        "mappings": [],
        "hasValidationIssue": "false",
    }
    next_id -= 1
    if email_field_id:
        step["keyFieldId"] = "100001"  # standard Email Address contact field id (common default)
        step["mappings"].append(
            {
                "type": "FormFieldUpdateMapping",
                "id": str(next_id),
                "sourceFormFieldId": email_field_id,
                "targetEntityFieldId": "100001",
                "updateType": "useFieldRule",
            }
        )
        next_id -= 1
    return step, next_id


def rows_to_form_json(
    rows: List[Dict[str, Any]],
    headers: List[str],
    form_name: str,
    form_description: str = "",
    processing_type: str = "externalWebsite",
    include_submit: bool = True,
    include_processing_step: bool = False,
) -> Dict[str, Any]:
    col_map = resolve_column_map(headers)
    if "name" not in col_map and "htmlName" not in col_map:
        raise ValueError(
            "Could not find a usable 'name' or 'htmlName' column.\n"
            f"Detected columns: {headers}\n"
            "Expected at least one of: name, field name, label, htmlName, html_name, ..."
        )

    elements: List[Dict[str, Any]] = []
    next_id = -1
    email_field_id: Optional[str] = None

    for row in rows:
        # Skip completely empty rows
        if all(
            (v is None or (isinstance(v, float) and pd is not None and pd.isna(v)) or str(v).strip() == "")
            for v in row.values()
        ):
            continue
        element, next_id = build_form_field(row, col_map, next_id)
        elements.append(element)
        # Heuristic: first email-ish field becomes the key field for processing step
        if email_field_id is None:
            hn = element.get("htmlName", "").lower()
            nm = element.get("name", "").lower()
            if "email" in hn or "email" in nm:
                email_field_id = element["id"]

    if not elements:
        raise ValueError("No field rows found in the file.")

    if include_submit:
        submit, next_id = build_submit_button(next_id)
        elements.append(submit)

    payload: Dict[str, Any] = {
        "type": "Form",
        "name": form_name.strip() or "Untitled Form",
        "elements": elements,
        "processingType": processing_type,
        "style": DEFAULT_FORM_STYLE,
    }
    if form_description.strip():
        payload["description"] = form_description.strip()

    if include_processing_step:
        step, next_id = build_basic_create_update_step(email_field_id, next_id)
        payload["processingSteps"] = [step]

    return payload


# ---------------------------------------------------------------------------
# GUI
# ---------------------------------------------------------------------------

class EloquaFormBuilderApp(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.title("Eloqua Form JSON Builder")
        self.geometry("1100x780")
        self.minsize(900, 650)

        self.dataframe: Optional[Any] = None
        self.file_path: Optional[str] = None
        self.headers: List[str] = []

        self._build_ui()
        self._set_status("Ready – load an Excel or CSV file of form fields.")

    # ----- UI construction -----

    def _build_ui(self) -> None:
        main = ttk.Frame(self, padding=8)
        main.pack(fill=tk.BOTH, expand=True)

        # --- Top: file + form options ---
        top = ttk.LabelFrame(main, text="1. Source & Form Options", padding=8)
        top.pack(fill=tk.X, pady=(0, 6))

        row0 = ttk.Frame(top)
        row0.pack(fill=tk.X, pady=2)
        ttk.Button(row0, text="Load Excel / CSV…", command=self.load_file).pack(side=tk.LEFT)
        self.file_label = ttk.Label(row0, text="No file loaded", foreground="#555")
        self.file_label.pack(side=tk.LEFT, padx=10)

        row1 = ttk.Frame(top)
        row1.pack(fill=tk.X, pady=4)
        ttk.Label(row1, text="Form name:").pack(side=tk.LEFT)
        self.form_name_var = tk.StringVar(value="My Eloqua Form")
        ttk.Entry(row1, textvariable=self.form_name_var, width=40).pack(side=tk.LEFT, padx=(4, 16))

        ttk.Label(row1, text="Processing type:").pack(side=tk.LEFT)
        self.proc_type_var = tk.StringVar(value="externalWebsite")
        ttk.Combobox(
            row1,
            textvariable=self.proc_type_var,
            values=["externalWebsite", "externalEmail", "internallyHosted", "internalAgentProxyForm"],
            width=22,
            state="readonly",
        ).pack(side=tk.LEFT, padx=4)

        row2 = ttk.Frame(top)
        row2.pack(fill=tk.X, pady=2)
        ttk.Label(row2, text="Description:").pack(side=tk.LEFT)
        self.form_desc_var = tk.StringVar()
        ttk.Entry(row2, textvariable=self.form_desc_var, width=70).pack(side=tk.LEFT, padx=4, fill=tk.X, expand=True)

        row3 = ttk.Frame(top)
        row3.pack(fill=tk.X, pady=4)
        self.include_submit_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(row3, text="Include Submit button", variable=self.include_submit_var).pack(side=tk.LEFT, padx=(0, 12))
        self.include_proc_step_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(
            row3,
            text="Add basic Create/Update Contact processing step",
            variable=self.include_proc_step_var,
        ).pack(side=tk.LEFT)

        ttk.Button(row3, text="Generate JSON from file", command=self.generate_json).pack(side=tk.RIGHT)

        # --- Middle: data preview ---
        mid = ttk.LabelFrame(main, text="2. Field Preview (from file)", padding=4)
        mid.pack(fill=tk.BOTH, expand=False, pady=(0, 6))

        tree_frame = ttk.Frame(mid)
        tree_frame.pack(fill=tk.BOTH, expand=True)
        self.tree = ttk.Treeview(tree_frame, show="headings", height=7)
        vsb = ttk.Scrollbar(tree_frame, orient=tk.VERTICAL, command=self.tree.yview)
        hsb = ttk.Scrollbar(tree_frame, orient=tk.HORIZONTAL, command=self.tree.xview)
        self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        self.tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")
        tree_frame.rowconfigure(0, weight=1)
        tree_frame.columnconfigure(0, weight=1)

        # --- Bottom: JSON editor ---
        bottom = ttk.LabelFrame(main, text="3. JSON Payload (review / edit)", padding=4)
        bottom.pack(fill=tk.BOTH, expand=True)

        toolbar = ttk.Frame(bottom)
        toolbar.pack(fill=tk.X, pady=(0, 4))
        ttk.Button(toolbar, text="Validate & Format", command=self.validate_and_format).pack(side=tk.LEFT, padx=(0, 6))
        ttk.Button(toolbar, text="Copy to Clipboard", command=self.copy_to_clipboard).pack(side=tk.LEFT, padx=(0, 6))
        ttk.Button(toolbar, text="Save JSON…", command=self.save_json).pack(side=tk.LEFT, padx=(0, 6))
        ttk.Button(toolbar, text="Clear", command=self.clear_json).pack(side=tk.LEFT)

        self.json_valid_var = tk.StringVar(value="")
        self.json_status = ttk.Label(toolbar, textvariable=self.json_valid_var, foreground="#0a7")
        self.json_status.pack(side=tk.RIGHT)

        self.json_text = scrolledtext.ScrolledText(
            bottom,
            wrap=tk.NONE,
            font=("Consolas", 11),
            undo=True,
            height=18,
        )
        self.json_text.pack(fill=tk.BOTH, expand=True)
        # Horizontal scroll for long lines
        xscroll = ttk.Scrollbar(bottom, orient=tk.HORIZONTAL, command=self.json_text.xview)
        self.json_text.configure(xscrollcommand=xscroll.set)
        xscroll.pack(fill=tk.X)

        # Status bar
        self.status_var = tk.StringVar()
        status = ttk.Label(self, textvariable=self.status_var, relief=tk.SUNKEN, anchor=tk.W, padding=(6, 2))
        status.pack(fill=tk.X, side=tk.BOTTOM)

        # Help hint
        hint = (
            "Expected columns (flexible names): name / label, htmlName, dataType, displayType, "
            "instructions, defaultValue, required, createdFromContactFieldId, fieldMergeId"
        )
        ttk.Label(main, text=hint, foreground="#666", font=("", 9)).pack(fill=tk.X, pady=(4, 0))

    # ----- Actions -----

    def _set_status(self, msg: str) -> None:
        self.status_var.set(msg)

    def load_file(self) -> None:
        if pd is None:
            messagebox.showerror(
                "Missing dependency",
                "pandas (and openpyxl for .xlsx) are required.\n\n"
                "Install with:\n  pip install pandas openpyxl",
            )
            return

        path = filedialog.askopenfilename(
            title="Select field definition file",
            filetypes=[
                ("Excel / CSV", "*.xlsx *.xls *.csv"),
                ("Excel", "*.xlsx *.xls"),
                ("CSV", "*.csv"),
                ("All files", "*.*"),
            ],
        )
        if not path:
            return

        try:
            if path.lower().endswith(".csv"):
                df = pd.read_csv(path, dtype=str, keep_default_na=False)
            else:
                df = pd.read_excel(path, dtype=str, keep_default_na=False)
        except Exception as exc:
            messagebox.showerror("Load error", f"Could not read file:\n{exc}")
            return

        # Drop fully empty columns / rows
        df = df.dropna(how="all", axis=0).dropna(how="all", axis=1)
        df.columns = [str(c).strip() for c in df.columns]

        self.dataframe = df
        self.file_path = path
        self.headers = list(df.columns)
        self.file_label.config(text=os.path.basename(path))

        # Populate tree
        self.tree.delete(*self.tree.get_children())
        self.tree["columns"] = self.headers
        for h in self.headers:
            self.tree.heading(h, text=h)
            self.tree.column(h, width=120, minwidth=60, stretch=True)
        for _, row in df.iterrows():
            values = [str(row[h]) if row[h] is not None else "" for h in self.headers]
            self.tree.insert("", tk.END, values=values)

        # Suggest form name from filename
        base = os.path.splitext(os.path.basename(path))[0]
        if not self.form_name_var.get().strip() or self.form_name_var.get() == "My Eloqua Form":
            self.form_name_var.set(base.replace("_", " ").replace("-", " ").title())

        self._set_status(f"Loaded {len(df)} row(s) from {os.path.basename(path)}")
        col_map = resolve_column_map(self.headers)
        mapped = ", ".join(f"{k}→{v}" for k, v in col_map.items()) or "(none auto-mapped)"
        self._set_status(self.status_var.get() + f"  |  Mapped: {mapped}")

    def generate_json(self) -> None:
        if self.dataframe is None or self.dataframe.empty:
            messagebox.showwarning("No data", "Load an Excel or CSV file first.")
            return

        try:
            rows = self.dataframe.to_dict(orient="records")
            payload = rows_to_form_json(
                rows=rows,
                headers=self.headers,
                form_name=self.form_name_var.get(),
                form_description=self.form_desc_var.get(),
                processing_type=self.proc_type_var.get(),
                include_submit=self.include_submit_var.get(),
                include_processing_step=self.include_proc_step_var.get(),
            )
        except Exception as exc:
            messagebox.showerror("Generation error", str(exc))
            return

        text = json.dumps(payload, indent=2, ensure_ascii=False)
        self.json_text.delete("1.0", tk.END)
        self.json_text.insert(tk.END, text)
        self.json_valid_var.set("✓ Valid JSON")
        self.json_status.configure(foreground="#0a7")
        self._set_status(f"Generated payload with {len(payload.get('elements', []))} element(s).")

    def validate_and_format(self) -> None:
        raw = self.json_text.get("1.0", tk.END).strip()
        if not raw:
            self.json_valid_var.set("Empty")
            self.json_status.configure(foreground="#a50")
            return
        try:
            obj = json.loads(raw)
            pretty = json.dumps(obj, indent=2, ensure_ascii=False)
            self.json_text.delete("1.0", tk.END)
            self.json_text.insert(tk.END, pretty)
            self.json_valid_var.set("✓ Valid JSON")
            self.json_status.configure(foreground="#0a7")
            self._set_status("JSON is valid and has been re-indented.")
        except json.JSONDecodeError as exc:
            self.json_valid_var.set(f"✗ Invalid: {exc.msg} (line {exc.lineno})")
            self.json_status.configure(foreground="#c00")
            self._set_status(f"JSON validation failed: {exc}")
            messagebox.showerror("Invalid JSON", f"{exc.msg}\nLine {exc.lineno}, column {exc.colno}")

    def copy_to_clipboard(self) -> None:
        raw = self.json_text.get("1.0", tk.END).strip()
        if not raw:
            messagebox.showwarning("Nothing to copy", "JSON editor is empty.")
            return
        # Optionally validate first
        try:
            json.loads(raw)
            valid = True
        except json.JSONDecodeError:
            valid = False
            if not messagebox.askyesno(
                "Invalid JSON",
                "The current text is not valid JSON.\nCopy anyway?",
            ):
                return

        self.clipboard_clear()
        self.clipboard_append(raw)
        self.update()  # keep clipboard after app focus changes
        msg = "Copied to clipboard."
        if not valid:
            msg += " (content is not valid JSON)"
        self._set_status(msg)
        self.json_valid_var.set("Copied ✓" if valid else "Copied (invalid)")
        self.json_status.configure(foreground="#0a7" if valid else "#a50")

    def save_json(self) -> None:
        raw = self.json_text.get("1.0", tk.END).strip()
        if not raw:
            messagebox.showwarning("Nothing to save", "JSON editor is empty.")
            return
        path = filedialog.asksaveasfilename(
            title="Save JSON",
            defaultextension=".json",
            filetypes=[("JSON", "*.json"), ("All files", "*.*")],
            initialfile=(self.form_name_var.get() or "eloqua_form").replace(" ", "_") + ".json",
        )
        if not path:
            return
        try:
            # Prefer writing the pretty version if valid
            try:
                obj = json.loads(raw)
                content = json.dumps(obj, indent=2, ensure_ascii=False)
            except json.JSONDecodeError:
                content = raw
            with open(path, "w", encoding="utf-8") as f:
                f.write(content)
            self._set_status(f"Saved: {path}")
        except Exception as exc:
            messagebox.showerror("Save error", str(exc))

    def clear_json(self) -> None:
        self.json_text.delete("1.0", tk.END)
        self.json_valid_var.set("")
        self._set_status("JSON cleared.")


def main() -> None:
    app = EloquaFormBuilderApp()
    app.mainloop()


if __name__ == "__main__":
    main()
