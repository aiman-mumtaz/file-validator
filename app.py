import csv
import io
import logging
import re
from datetime import datetime
from typing import Dict, Iterable, List, Tuple

import streamlit as st
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment

logger = logging.getLogger(__name__)

OVERPUNCH_POSITIVE = {
'{': '0', 'A': '1', 'B': '2', 'C': '3', 'D': '4',
'E': '5', 'F': '6', 'G': '7', 'H': '8', 'I': '9'
}
OVERPUNCH_NEGATIVE = {
'}': '0', 'J': '1', 'K': '2', 'L': '3', 'M': '4',
'N': '5', 'O': '6', 'P': '7', 'Q': '8', 'R': '9'
}


class ValidationError(Exception):
pass


def _normalize_header(value: str) -> str:
return re.sub(r'[^a-z0-9]', '', value.strip().lower()) if value else ''


def _parse_bool(value: str) -> bool:
if value is None:
return False
return str(value).strip().lower() in {'true', 'yes', 'y', '1'}


def _parse_master_seq(value: str) -> int | None:
if value is None:
return None
cleaned = str(value).strip()
if not cleaned:
return None
try:
return int(float(cleaned))
except ValueError:
return None


def _load_field_mapping(mapping_stream: io.TextIOBase) -> Dict[str, dict]:
field_mapping: Dict[str, dict] = {}
reader = csv.DictReader(mapping_stream)
header_map = { _normalize_header(name): name for name in (reader.fieldnames or []) }
logger.info("Mapping headers detected: %s", ", ".join(reader.fieldnames or []))

def get_field(row: dict, *candidates: str) -> str | None:
for candidate in candidates:
key = header_map.get(_normalize_header(candidate))
if key and key in row:
return row.get(key)
for candidate in candidates:
if candidate in row:
return row.get(candidate)
return None

for row in reader:
try:
master_seq_val = _parse_master_seq(get_field(row, 'Master Seq No', 'MasterSeqNo', 'Master Sequence No'))
if master_seq_val != 4:
continue

field_name = get_field(row, 'Field Name', 'FieldName')
field_format = get_field(row, 'Field Format', 'FieldFormat')
field_length_raw = get_field(row, 'Field Length', 'FieldLength')
field_length = int(field_length_raw) if field_length_raw else 0

is_numerical = _parse_bool(get_field(row, 'IsNumericalFormat', 'IsNumerical'))
is_overpunch = _parse_bool(get_field(row, 'IsOverpunchFormat', 'IsOverpunch'))
format_spec = get_field(row, 'Format', 'FieldFormatSpec')
if format_spec and str(format_spec).strip().lower() == 'null':
format_spec = None

start_pos = get_field(row, 'start', 'Start', 'Start Position', 'StartPosition')
end_pos = get_field(row, 'end', 'End', 'End Position', 'EndPosition')

apply_alpha_pad = get_field(row, 'applyAlphaPad', 'ApplyAlphaPad') or ''
apply_alpha_len = get_field(row, 'applyAlphaLen', 'ApplyAlphaLen') or ''
apply_numeric_pad = get_field(row, 'applyNumericPad', 'ApplyNumericPad') or ''
apply_numeric_len = get_field(row, 'applyNumericLen', 'ApplyNumericLen') or ''

if field_length > 0 and field_name and start_pos and end_pos:
decimal_places = 0
if format_spec and 'V' in str(format_spec):
match = re.search(r'V9+', str(format_spec))
if match:
decimal_places = len(match.group()) - 1

alpha_pad_len = int(apply_alpha_len) if str(apply_alpha_len).isdigit() else 0
numeric_pad_len = int(apply_numeric_len) if str(apply_numeric_len).isdigit() else 0

field_mapping[str(field_name).strip()] = {
'format': field_format,
'length': field_length,
'is_numerical': is_numerical,
'is_overpunch': is_overpunch,
'decimal_places': decimal_places,
'format_spec': format_spec,
'start': int(float(start_pos)) - 1,
'end': int(float(end_pos)),
'alpha_pad_field': apply_alpha_pad,
'alpha_pad_len': alpha_pad_len,
'numeric_pad_field': apply_numeric_pad,
'numeric_pad_len': numeric_pad_len
}
except (ValueError, KeyError, TypeError):
continue

if not field_mapping:
found_headers = ', '.join(reader.fieldnames or [])
logger.error("No field mappings found in CSV. Headers: %s", found_headers or 'none')
raise ValidationError(
"No field mappings found in CSV. Verify the mapping file format. "
f"Found headers: {found_headers or 'none'}"
)

return field_mapping


def decode_overpunch(value: str) -> str:
if not value:
return value

last_char = value[-1]
if last_char in OVERPUNCH_POSITIVE:
return value[:-1] + OVERPUNCH_POSITIVE[last_char]
if last_char in OVERPUNCH_NEGATIVE:
return '-' + value[:-1] + OVERPUNCH_NEGATIVE[last_char]
return value


def normalize_field(raw_value: str, field_info: dict) -> str:
if not raw_value:
return ''

cleaned = ''.join(char if char.isprintable() else ' ' for char in raw_value).strip()

if not cleaned:
return ''

is_numerical = field_info.get('is_numerical', False)
is_overpunch = field_info.get('is_overpunch', False)
decimal_places = field_info.get('decimal_places', 0)

if is_overpunch:
cleaned = decode_overpunch(cleaned)

if is_numerical:
try:
num_str = ''.join(c for c in cleaned if c.isdigit() or c == '-')
if not num_str or num_str == '-' or num_str.replace('-', '') == '0' * len(num_str.replace('-', '')):
return ''

is_negative = num_str.startswith('-')
num_str = num_str.replace('-', '')

if decimal_places > 0 and num_str:
if len(num_str) > decimal_places:
integer_part = num_str[:-decimal_places]
fractional_part = num_str[-decimal_places:]
num_val = float(f"{integer_part}.{fractional_part}")
else:
num_val = float(f"0.{num_str.zfill(decimal_places)}")

if is_negative:
num_val = -num_val

result = str(num_val)
if '.' in result:
result = result.rstrip('0').rstrip('.')
return result

num_val = int(num_str)
if is_negative:
num_val = -num_val
return str(num_val)
except Exception:
pass

return cleaned


def parse_claim_with_mapping(line: str, field_mapping: Dict[str, dict]) -> dict:
if not line.startswith('4'):
return None

claim = {}
claim_raw = {}
claim_raw_exact = {}
