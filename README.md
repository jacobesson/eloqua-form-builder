# Eloqua Form JSON Builder (Tkinter)

Desktop helper for building the JSON body required by:

```
POST /api/rest/2.0/assets/form
```

Docs: https://docs.oracle.com/en/cloud/saas/marketing/eloqua-rest-api/op-api-rest-2.0-assets-form-post.html

## Requirements

```bash
pip install pandas openpyxl
```

(Tkinter is included with standard CPython on Windows/macOS; on some Linux distros install `python3-tk`.)

## Run

```bash
python eloqua_form_builder.py
```

## Workflow

1. **Load Excel / CSV** – one row per form field.
2. Set **Form name**, optional description, processing type.
3. Optionally keep the Submit button and/or a basic Create/Update Contact processing step.
4. Click **Generate JSON from file**.
5. Review / edit the JSON in the editor.
6. **Validate & Format** (pretty-prints with correct indentation).
7. **Copy to Clipboard** or **Save JSON**.

## Expected columns (flexible names)

| Canonical        | Accepted aliases                                      |
|------------------|-------------------------------------------------------|
| name             | name, field name, label, display name                 |
| htmlName         | htmlName, html_name, field_name, api name             |
| dataType         | dataType, data_type, type                             |
| displayType      | displayType, display_type, ui type, control           |
| instructions     | instructions, help text, description, hint            |
| defaultValue     | defaultValue, default_value, default                  |
| required         | required, is_required, mandatory (true/yes/1)         |
| createdFromContactFieldId | contact field id variants                    |
| fieldMergeId     | field merge id variants                               |

A sample file `sample_form_fields.csv` is included.

## Notes on Eloqua payload

- New form fields must use **negative temporary IDs** (`-1`, `-2`, …). The app assigns them automatically.
- A Submit button (`displayType: "submit"`) is added by default.
- The optional processing step uses the common contact Email field id `100001` as key field when an email-like column is detected.
- You can freely edit the generated JSON; validation only checks that it is well-formed JSON (not a full Eloqua schema validator).
