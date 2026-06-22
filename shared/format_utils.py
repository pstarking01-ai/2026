import re


def format_comma(all_fields, group, field):
    """Add comma separators to numeric input in real time."""
    edit = all_fields[group][field]
    text = edit.text().replace(",", "")
    if text.isdigit():
        edit.blockSignals(True)
        try:
            edit.setText(format(int(text), ","))
        except Exception:
            pass
        edit.blockSignals(False)


def format_phone(all_fields, group, field_name):
    """Auto-insert hyphens into phone number input."""
    edit = all_fields[group][field_name]
    text = re.sub(r'[^0-9]', '', edit.text())
    if len(text) >= 10:
        formatted = f"{text[:3]}-{text[3:7]}-{text[7:]}"
    elif len(text) > 6:
        formatted = f"{text[:3]}-{text[3:6]}-{text[6:]}"
    else:
        formatted = text
    edit.blockSignals(True)
    edit.setText(formatted)
    edit.blockSignals(False)
