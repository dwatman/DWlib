import os
import re
import csv
import shutil
import argparse
from datetime import datetime

# --- CONFIGURATION ---
MASTER_CSV_NAME = "master_parts.csv"
FIELDS_TO_EXTRACT = {
	"MPN": "MPN",
	"Mouser": "Mouser PN",
	"Digikey": "Digikey PN",
	"Manufacturer": "Manufacturer",
	"Tolerance": "Tolerance",
	"Extra": "Extra",
	"Description": "Description"
}

# Component types where we should parse and normalize values
VALUE_PARSE_TYPES = {'R', 'C', 'L', 'D', 'BT', 'F', 'FB', 'Y', 'Z'}

def clean_field_value(value):
	"""Treat ~ or - as empty field values (KiCad placeholders)"""
	cleaned = value.strip()
	if cleaned in ['~', '-']:
		return ""
	return cleaned

def extract_voltage(value_str):
	"""Extract voltage from value string (e.g., '100n 16V' -> ('100n', '16V'))"""
	if not value_str:
		return value_str, ""

	# Look for voltage patterns: number followed by V (possibly with decimal)
	voltage_pattern = r'\s*(\d+\.?\d*V)\s*'
	match = re.search(voltage_pattern, value_str)

	if match:
		voltage = match.group(1)
		# Remove the voltage from the value string
		cleaned_value = re.sub(voltage_pattern, ' ', value_str).strip()
		# Clean up multiple spaces
		cleaned_value = re.sub(r'\s+', ' ', cleaned_value)
		return cleaned_value, voltage

	return value_str, ""

def normalize_value(value_str, component_type):
	"""Normalize component value to consistent format (e.g., 4k7 -> 4.7k)"""
	if not value_str or component_type not in VALUE_PARSE_TYPES:
		return value_str

	# First, clean up extra spaces
	value_str = re.sub(r'\s+', ' ', value_str.strip())

	# Split off any suffix text (like X7R, 1%, etc.)
	parts = value_str.split()
	if not parts:
		return value_str

	numeric_part = parts[0]
	suffix = ' '.join(parts[1:]) if len(parts) > 1 else ""

	# Define multipliers
	multipliers = {
		'p': (1e-12, 'p'),
		'n': (1e-9, 'n'),
		'u': (1e-6, 'u'),
		'µ': (1e-6, 'u'),  # Convert µ to u for consistency
		'm': (1e-3, 'm'),
		'k': (1e3, 'k'),
		'K': (1e3, 'k'),
		'M': (1e6, 'M'),
		'G': (1e9, 'G')
	}

	# Units to strip
	units_to_strip = ['Ω', 'Ohm', 'ohms', 'ohm', 'F', 'uF', 'nF', 'pF', 'H', 'uH', 'nH', 'mH', 'Hz', 'kHz', 'MHz']

	# Remove units from the numeric part
	working_str = numeric_part
	for unit in units_to_strip:
		working_str = re.sub(re.escape(unit) + r'$', '', working_str, flags=re.IGNORECASE)

	# Try to parse and normalize
	# Format 1: 10k, 4.7u, etc. (number followed by multiplier)
	match1 = re.match(r'^(\d+\.?\d*)([pnuµmkKMG])$', working_str)
	if match1:
		num, mult = match1.groups()
		if mult == 'µ':
			mult = 'u'
		elif mult == 'K':
			mult = 'k'
		normalized = f"{num}{mult}"
		return f"{normalized} {suffix}".strip() if suffix else normalized

	# Format 2: 4k7, 1M2, etc. (number, multiplier, number)
	match2 = re.match(r'^(\d+)([pnuµmkKMG])(\d+)$', working_str)
	if match2:
		main, mult, dec = match2.groups()
		if mult == 'µ':
			mult = 'u'
		elif mult == 'K':
			mult = 'k'
		# Convert to decimal format (4k7 -> 4.7k)
		decimal_val = f"{main}.{dec}"
		normalized = f"{decimal_val}{mult}"
		return f"{normalized} {suffix}".strip() if suffix else normalized

	# Format 3: Special resistor format like 5R or 2R2 (5 ohms, 2.2 ohms)
	# Keep the R since it indicates the unit when there's no multiplier
	match3 = re.match(r'^(\d+)R(\d*)$', working_str, re.IGNORECASE)
	if match3:
		main, dec = match3.groups()
		if dec:
			normalized = f"{main}.{dec}R"
		else:
			normalized = f"{main}R"
		return f"{normalized} {suffix}".strip() if suffix else normalized

	# Format 4: Just a number (no multiplier)
	match4 = re.match(r'^(\d+\.?\d*)$', working_str)
	if match4:
		return f"{working_str} {suffix}".strip() if suffix else working_str

	# If we can't parse it, return original
	return value_str

def parse_value_to_float(value_str):
	"""Parses component value (e.g. 10k, 4.7u) to float for sorting."""
	if not value_str:
		return 0.0

	# Take just the first part (before any space)
	s = value_str.split()[0]

	# Define multipliers with their numeric values
	multipliers = {
		'p': 1e-12,
		'n': 1e-9,
		'u': 1e-6,
		'µ': 1e-6,
		'm': 1e-3,
		'k': 1e3,
		'K': 1e3,
		'M': 1e6,
		'G': 1e9
	}

	# Try format: "10k" or "4.7u"
	match = re.match(r'^([0-9\.]+)([pnuµmkKMG]?)$', s)
	if match:
		num, mult = match.groups()
		try:
			val = float(num)
			if mult and mult in multipliers:
				val *= multipliers[mult]
			return val
		except ValueError:
			pass

	# Try format: "4k7" or "1M2"
	match_mid = re.match(r'^(\d+)([pnuµmkKMG])(\d+)$', s)
	if match_mid:
		main, mult, dec = match_mid.groups()
		try:
			val = float(f"{main}.{dec}")
			if mult in multipliers:
				val *= multipliers[mult]
			return val
		except ValueError:
			pass

	# Try format: "5R" or "2R2" (resistor format)
	match_r = re.match(r'^(\d+)R(\d*)$', s, re.IGNORECASE)
	if match_r:
		main, dec = match_r.groups()
		try:
			if dec:
				return float(f"{main}.{dec}")
			else:
				return float(main)
		except ValueError:
			pass

	return 0.0

def parse_voltage_to_float(volt_str):
	"""Parses voltage string (e.g. 16V, 6.3V) to float for sorting."""
	if not volt_str:
		return 0.0
	# Extract the first number found
	match = re.search(r'([0-9\.]+)', volt_str)
	if match:
		try:
			return float(match.group(1))
		except ValueError:
			return 0.0
	return 0.0

def get_designator_type(refdes):
	"""Extract the letter prefix from reference designator (e.g., U1A -> U)"""
	match = re.match(r'([A-Za-z]+)', refdes)
	return match.group(1).upper() if match else "ZZ"

def strip_designator_suffix(refdes):
	"""Remove unit suffixes from reference designator (e.g., U1A -> U1)"""
	# Match letter prefix + number, ignore any trailing letters/numbers
	match = re.match(r'([A-Za-z]+\d+)', refdes)
	return match.group(1) if match else refdes

def get_balanced_block(text, start_index):
	"""Find the closing parenthesis for an S-expression block"""
	depth = 0
	in_string = False
	for i in range(start_index, len(text)):
		char = text[i]
		if char == '"' and (i == 0 or text[i-1] != '\\'):
			in_string = not in_string
			continue
		if in_string:
			continue
		if char == '(':
			depth += 1
		elif char == ')':
			depth -= 1
			if depth == 0:
				return i
	return len(text)

def should_exclude_component(block_text):
	"""Check if component should be excluded (DNP, not in BOM, not on board)"""
	# Check for (dnp yes)
	if re.search(r'\(dnp\s+yes\)', block_text):
		return True

	# Check for (in_bom no)
	if re.search(r'\(in_bom\s+no\)', block_text):
		return True

	# Check for (on_board no)
	if re.search(r'\(on_board\s+no\)', block_text):
		return True

	return False

def extract_lib_id(block_text):
	"""Extract the lib_id (symbol) from the component block"""
	match = re.search(r'\(lib_id\s+"([^"]+)"', block_text)
	return match.group(1) if match else ""

def parse_schematic_file(filepath):
	"""Parses KiCad 9 S-expressions, skipping library definitions."""
	with open(filepath, 'r', encoding='utf-8') as f:
		content = f.read()

	components = {}

	# Remove (lib_symbols ...) block
	lib_sym_start = content.find('(lib_symbols')
	if lib_sym_start != -1:
		lib_sym_end = get_balanced_block(content, lib_sym_start)
		content = content[:lib_sym_start] + content[lib_sym_end+1:]

	# Find instance blocks
	symbol_indices = [m.start() for m in re.finditer(r'\(symbol', content)]

	for start_idx in symbol_indices:
		end_idx = get_balanced_block(content, start_idx)
		block_text = content[start_idx : end_idx+1]

		if '(lib_id' not in block_text:
			continue

		# Check if component should be excluded
		if should_exclude_component(block_text):
			continue

		props_list = re.findall(r'\(property\s+"([^"]+)"\s+"((?:[^"\\]|\\.)*)"', block_text)
		props = {k: v for k, v in props_list}

		refdes = props.get("Reference", "")
		if not refdes:
			continue

		# Skip reference designators starting with # (power symbols, etc.)
		if refdes.startswith("#"):
			continue

		# Extract lib_id (symbol)
		lib_id = extract_lib_id(block_text)
		props["_lib_id"] = lib_id

		# Deduplicate multi-unit parts (U1A, U1B -> U1)
		base_refdes = strip_designator_suffix(refdes)
		if base_refdes not in components:
			components[base_refdes] = props

	return list(components.values())

def scan_project_folder(folder_path):
	project_name = os.path.basename(os.path.normpath(folder_path))
	extracted_parts = {}

	# Track the newest file date in the project
	latest_mod_time = 0.0

	sch_files = [f for f in os.listdir(folder_path) if f.endswith('.kicad_sch')]
	print(f"Scanning project '{project_name}' ({len(sch_files)} sheets found)...")

	for sch_file in sch_files:
		path = os.path.join(folder_path, sch_file)

		# Check file modification time
		mtime = os.path.getmtime(path)
		if mtime > latest_mod_time:
			latest_mod_time = mtime

		found_components = parse_schematic_file(path)

		for props in found_components:
			refdes = props.get("Reference", "")
			mpn = clean_field_value(props.get(FIELDS_TO_EXTRACT["MPN"], ""))

			if not mpn:
				continue

			val = clean_field_value(props.get("Value", ""))
			voltage = clean_field_value(props.get("Voltage", ""))
			footprint = clean_field_value(props.get("Footprint", ""))
			mouser = clean_field_value(props.get(FIELDS_TO_EXTRACT["Mouser"], ""))
			digikey = clean_field_value(props.get(FIELDS_TO_EXTRACT["Digikey"], ""))
			manufacturer = clean_field_value(props.get(FIELDS_TO_EXTRACT["Manufacturer"], ""))
			tolerance = clean_field_value(props.get(FIELDS_TO_EXTRACT["Tolerance"], ""))
			extra = clean_field_value(props.get(FIELDS_TO_EXTRACT["Extra"], ""))

			# Extract new fields
			description = clean_field_value(props.get(FIELDS_TO_EXTRACT["Description"], ""))
			symbol = clean_field_value(props.get("_lib_id", ""))
			datasheet = clean_field_value(props.get("Datasheet", ""))

			# Skip parts with no valid footprint
			if not footprint:
				continue

			# Get component type
			comp_type = get_designator_type(refdes)

			# Process value and voltage for applicable component types
			if comp_type in VALUE_PARSE_TYPES:
				# If Voltage property is not set, try to extract from Value
				if not voltage and val:
					val, extracted_voltage = extract_voltage(val)
					if extracted_voltage:
						voltage = extracted_voltage
				elif voltage and val:
					# Voltage property exists, but still clean Value field
					val, _ = extract_voltage(val)

				# Normalize the value format
				val = normalize_value(val, comp_type)

			# Format date from file timestamp
			date_str = datetime.fromtimestamp(latest_mod_time).strftime("%Y-%m-%d") if latest_mod_time > 0 else datetime.now().strftime("%Y-%m-%d")

			part_data = {
				"Type": comp_type,
				"Value": val,
				"Voltage": voltage,
				"Tolerance": tolerance,
				"Extra": extra,
				"Footprint": footprint,
				"Stock": "",  # Empty, for manual entry
				"Manufacturer": manufacturer,
				"MPN": mpn,
				"Digikey PN": digikey,
				"Mouser PN": mouser,
				"Description": description,
				"Symbol": symbol,
				"Datasheet": datasheet,
				"Projects": {project_name},
				"Last Used": date_str
			}

			if mpn in extracted_parts:
				existing = extracted_parts[mpn]
				# Consistency Check
				if existing['Value'] != val or existing['Footprint'] != footprint:
					print(f"WARNING: Consistency Check Failed within project for MPN {mpn}!")
					print(f"  > OLD ({existing['Type']}): {existing['Value']} [{existing['Footprint']}]")
					print(f"  > NEW ({refdes}): {val} [{footprint}]")
			else:
				extracted_parts[mpn] = part_data

	return extracted_parts

def load_master_csv(filepath):
	parts = {}
	if not os.path.exists(filepath):
		return parts

	with open(filepath, 'r', newline='', encoding='utf-8') as csvfile:
		reader = csv.DictReader(csvfile)
		for row in reader:
			mpn = row['MPN']
			projs = set(row['Used In'].split(', ')) if row['Used In'] else set()

			parts[mpn] = {
				"Type": row['Type'],
				"Value": row['Value'],
				"Voltage": row.get('Voltage', ''),
				"Tolerance": row.get('Tolerance', ''),
				"Extra": row.get('Extra', ''),
				"Footprint": row['Footprint'],
				"Stock": row.get('Stock', ''),
				"Manufacturer": row.get('Manufacturer', ''),
				"MPN": mpn,
				"Digikey PN": row['Digikey PN'],
				"Mouser PN": row['Mouser PN'],
				"Description": row.get('Description', ''),
				"Symbol": row.get('Symbol', ''),
				"Datasheet": row.get('Datasheet', ''),
				"Projects": projs,
				"Last Used": row['Last Used']
			}
	return parts

def merge_data(master_data, new_data):
	changes = []
	errors = []

	for mpn, new_part in new_data.items():
		if mpn in master_data:
			existing = master_data[mpn]
			# Consistency Check
			if existing['Value'] != new_part['Value'] or existing['Footprint'] != new_part['Footprint']:
				errors.append(f"CONFLICT: MPN {mpn} exists but specs differ!\n   Master: {existing['Value']} [{existing['Footprint']}]\n   New:    {new_part['Value']} [{new_part['Footprint']}]")
				continue

			updated = False
			# Update blank fields (but preserve Stock manually entered)
			if not existing['Mouser PN'] and new_part['Mouser PN']:
				existing['Mouser PN'] = new_part['Mouser PN']
				updated = True
			if not existing['Digikey PN'] and new_part['Digikey PN']:
				existing['Digikey PN'] = new_part['Digikey PN']
				updated = True
			if not existing['Voltage'] and new_part['Voltage']:
				existing['Voltage'] = new_part['Voltage']
				updated = True
			if not existing['Tolerance'] and new_part['Tolerance']:
				existing['Tolerance'] = new_part['Tolerance']
				updated = True
			if not existing['Manufacturer'] and new_part['Manufacturer']:
				existing['Manufacturer'] = new_part['Manufacturer']
				updated = True
			if not existing['Datasheet'] and new_part['Datasheet']:
				existing['Datasheet'] = new_part['Datasheet']
				updated = True
			if not existing['Extra'] and new_part['Extra']:
				existing['Extra'] = new_part['Extra']
				updated = True

			# Always update Description and Symbol to newest (with warning if changed)
			if existing['Description'] != new_part['Description'] and new_part['Description']:
				if existing['Description']:
					print(f"INFO: Updating description for {mpn}")
					print(f"  > Old: {existing['Description']}")
					print(f"  > New: {new_part['Description']}")
				existing['Description'] = new_part['Description']
				updated = True
			elif not existing['Description'] and new_part['Description']:
				existing['Description'] = new_part['Description']
				updated = True

			if existing['Symbol'] != new_part['Symbol'] and new_part['Symbol']:
				if existing['Symbol']:
					print(f"INFO: Updating symbol for {mpn}")
					print(f"  > Old: {existing['Symbol']}")
					print(f"  > New: {new_part['Symbol']}")
				existing['Symbol'] = new_part['Symbol']
				updated = True
			elif not existing['Symbol'] and new_part['Symbol']:
				existing['Symbol'] = new_part['Symbol']
				updated = True

			new_proj = list(new_part['Projects'])[0]
			if new_proj not in existing['Projects']:
				existing['Projects'].add(new_proj)
				# Always take the NEWER date
				if new_part['Last Used'] > existing['Last Used']:
					existing['Last Used'] = new_part['Last Used']
				updated = True
				changes.append(f"Updated {mpn}: Added project '{new_proj}'")
			elif updated:
				changes.append(f"Updated {mpn}: Filled missing data")
		else:
			master_data[mpn] = new_part
			changes.append(f"New Part: {new_part['Type']} {new_part['Value']} ({mpn})")

	return changes, errors

def main():
	parser = argparse.ArgumentParser(description="Update KiCad Master Parts List")
	parser.add_argument("project_path", help="Path to the KiCad project folder")
	args = parser.parse_args()

	folder_path = args.project_path
	if not os.path.isdir(folder_path):
		print("Error: Invalid directory path.")
		return

	master_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), MASTER_CSV_NAME)
	master_data = load_master_csv(master_path)
	new_data = scan_project_folder(folder_path)

	if not new_data:
		print("No parts with MPNs found in this project.")
		return

	changes, errors = merge_data(master_data, new_data)

	print("\n" + "="*40 + "\n SCAN RESULTS\n" + "="*40)
	if errors:
		print(f"\nFound {len(errors)} CONFLICTS (Will not be updated):")
		for e in errors:
			print(f"- {e}")

	if changes:
		print(f"\nFound {len(changes)} updates/additions:")
		for c in changes:
			print(f"- {c}")
	else:
		print("\nNo changes found. Master list is up to date.")
		return

	confirm = input("\nSave these changes to master_parts.csv? (y/n): ")
	if confirm.lower() != 'y':
		return

	if os.path.exists(master_path):
		backup_name = f"{MASTER_CSV_NAME}.bak"
		shutil.copy(master_path, os.path.join(os.path.dirname(master_path), backup_name))
		print(f"Backup created: {backup_name}")

	# Sorting Logic: Type -> Value -> Voltage -> Footprint -> MPN
	# Use numeric sorting for component types with parseable values, alphabetic for others
	sorted_parts = sorted(
		master_data.values(),
		key=lambda x: (
			x['Type'],
			parse_value_to_float(x['Value']) if x['Type'] in VALUE_PARSE_TYPES else 0.0,
			x['Value'] if x['Type'] not in VALUE_PARSE_TYPES else "",
			parse_voltage_to_float(x['Voltage']),
			x['Footprint'],
			x['MPN']
		)
	)

	# Column order optimized for visibility and workflow
	fieldnames = [
		"Type", "Value", "Voltage", "Tolerance", "Extra", "Footprint",
		"Stock",
		"Manufacturer", "MPN", "Digikey PN", "Mouser PN",
		"Description", "Symbol", "Datasheet",
		"Last Used", "Used In"
	]

	try:
		with open(master_path, 'w', newline='', encoding='utf-8') as csvfile:
			writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
			writer.writeheader()
			for part in sorted_parts:
				row_data = part.copy()
				row_data['Used In'] = ", ".join(sorted(part['Projects']))
				del row_data['Projects']
				writer.writerow(row_data)
		print("Success! Master list updated.")
	except PermissionError:
		print("ERROR: Is 'master_parts.csv' open? Close it and try again.")

if __name__ == "__main__":
	main()
