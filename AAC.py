import streamlit as st
import pandas as pd
import re
import os
import gdown
from io import BytesIO

# For updating file on Drive
from pydrive2.auth import GoogleAuth
from pydrive2.drive import GoogleDrive

# ------------------ Global Constants & Variables ------------------ #
MAPPING_FILE_ID = "1QP1XnxyDEgfxYfgBg_mf2ngXNfm9O8s5"  # your Google Sheets file ID
MULTIPLIER_MAPPING = {
    'k': 1e3,    # kilo
    'M': 1e6,    # mega
    'G': 1e9,    # giga
    'T': 1e12,   # tera
    'P': 1e15,   # peta
    'E': 1e18,   # exa
    'Z': 1e21,   # zetta
    'Y': 1e24,   # yotta
    'd': 1e-1,   # deci
    'c': 1e-2,   # centi
    'm': 1e-3,   # milli
    'µ': 1e-6,   # micro
    'n': 1e-9,   # nano
    'p': 1e-12,  # pico
    'f': 1e-15,  # femto
    'a': 1e-18,  # atto
    'z': 1e-21,  # zepto
    'y': 1e-24   # yocto
}

# ------------------ Helper Functions ------------------ #
def download_mapping_file():
    # Export the Google Sheet as an Excel file using the export endpoint
    mapping_url = f"https://docs.google.com/spreadsheets/d/{MAPPING_FILE_ID}/export?format=xlsx"
    output_path = "mapping.xlsx"
    if not os.path.exists(output_path):
        st.info("Downloading mapping file from Google Drive...")
        gdown.download(mapping_url, output_path, quiet=False)
    return output_path

def read_mapping_file(mapping_file_path):
    if not os.path.exists(mapping_file_path):
        raise FileNotFoundError(f"Error: '{mapping_file_path}' not found.")
    try:
        mapping_df = pd.read_excel(mapping_file_path)
    except Exception as e:
        raise Exception(f"Error reading '{mapping_file_path}': {e}")
    required_columns = {'Base Unit Symbol', 'Multiplier Symbol'}
    if not required_columns.issubset(mapping_df.columns):
        raise ValueError(f"'{mapping_file_path}' must contain the columns: {required_columns}")
    base_units = {str(unit).strip() for unit in mapping_df['Base Unit Symbol'].dropna().unique()}
    # Check for undefined multipliers (skip rows where multiplier is null)
    multipliers_df = mapping_df[mapping_df['Multiplier Symbol'].notna()]
    defined_multipliers = set(multipliers_df['Multiplier Symbol'])
    undefined_multipliers = defined_multipliers - set(MULTIPLIER_MAPPING.keys())
    if undefined_multipliers:
        raise ValueError(f"Undefined multipliers found in '{mapping_file_path}': {undefined_multipliers}")
    return mapping_df, base_units, MULTIPLIER_MAPPING

def split_outside_parens(text, delimiters):
    tokens = []
    current = ""
    i = 0
    depth = 0
    sorted_delims = sorted(delimiters, key=len, reverse=True)
    while i < len(text):
        ch = text[i]
        if ch == '(':
            depth += 1
            current += ch
            i += 1
        elif ch == ')':
            depth = max(depth - 1, 0)
            current += ch
            i += 1
        elif depth == 0:
            matched = None
            for delim in sorted_delims:
                if text[i:i+len(delim)] == delim:
                    matched = delim
                    break
            if matched:
                if current:
                    tokens.append(current)
                tokens.append(matched)
                current = ""
                i += len(matched)
            else:
                current += ch
                i += 1
        else:
            current += ch
            i += 1
    if current:
        tokens.append(current)
    return tokens

def process_unit_token_no_paren(token, base_units, multipliers_dict):
    if token.startswith('$'):
        after_dollar_preserved = token[1:]
        after_dollar_stripped = after_dollar_preserved.strip()
        if after_dollar_stripped == "":
            return "$"
        if after_dollar_stripped in base_units:
            return "$" + after_dollar_preserved
        sorted_prefixes = sorted(multipliers_dict.keys(), key=len, reverse=True)
        for prefix in sorted_prefixes:
            if after_dollar_stripped.startswith(prefix):
                possible_base_stripped = after_dollar_stripped[len(prefix):]
                if possible_base_stripped in base_units:
                    idx = after_dollar_preserved.find(prefix)
                    if idx != -1:
                        if idx == 1 and after_dollar_preserved[0] == " ":
                            base_unit_preserved = after_dollar_preserved[:0] + after_dollar_preserved[idx + len(prefix):]
                        else:
                            base_unit_preserved = after_dollar_preserved[:idx] + after_dollar_preserved[idx + len(prefix):]
                    else:
                        base_unit_preserved = possible_base_stripped
                    return "$" + base_unit_preserved
        return f"Error: Undefined unit '{after_dollar_stripped}' (no recognized prefix)"
    else:
        stripped_token = token.strip()
        if stripped_token in base_units:
            return "$" + stripped_token
        sorted_prefixes = sorted(multipliers_dict.keys(), key=len, reverse=True)
        for prefix in sorted_prefixes:
            if stripped_token.startswith(prefix):
                possible_base_stripped = stripped_token[len(prefix):]
                if possible_base_stripped in base_units:
                    idx = token.find(prefix)
                    if idx != -1:
                        if idx > 0 and token[idx-1] == " ":
                            if idx == 1:
                                base_unit_preserved = token[:0] + token[idx + len(prefix):]
                            else:
                                base_unit_preserved = token[:idx] + token[idx + len(prefix):]
                        else:
                            base_unit_preserved = token[:idx] + token[idx + len(prefix):]
                    else:
                        base_unit_preserved = possible_base_stripped
                    return "$" + base_unit_preserved
        return f"Error: Undefined unit '{stripped_token}' (no recognized prefix)"

def process_unit_token(token, base_units, multipliers_dict):
    pattern = re.compile(
        r'^(?P<lead>\s*)'
        r'(?P<numeric>[+\-±]?\d*(?:\.\d+)?)(?P<space1>\s*)'
        r'(?P<unit>.*?)(?P<space2>\s*)'
        r'(?P<paren>\([^)]*\))?'
        r'(?P<trail>\s*)$'
    )
    m = pattern.match(token)
    if not m:
        return token
    lead     = m.group('lead')
    numeric  = m.group('numeric')
    space1   = m.group('space1')
    unit_part = m.group('unit')
    space2   = m.group('space2')
    paren    = m.group('paren') if m.group('paren') else ""
    trail    = m.group('trail')
    core = unit_part.strip()
    left_ws  = re.match(r'^\s*', unit_part).group(0) or ""
    right_ws = re.search(r'\s*$', unit_part).group(0) or ""
    processed_core = process_unit_token_no_paren(core, base_units, multipliers_dict)
    new_unit_part = left_ws + processed_core + right_ws
    if "ohm" in core.lower():
        if new_unit_part.startswith("$") and not new_unit_part.startswith("$ "):
            new_unit_part = "$ " + new_unit_part[1:].lstrip()
        if numeric and not space1:
            space1 = " "
    return f"{lead}{numeric}{space1}{new_unit_part}{space2}{paren}{trail}"

def resolve_compound_unit(normalized_unit, base_units, multipliers_dict):
    tokens = split_outside_parens(normalized_unit, delimiters=["to", ",", "@"])
    resolved_parts = []
    for part in tokens:
        if part in ["to", ",", "@"]:
            resolved_parts.append(part)
        else:
            if part == "":
                continue
            resolved_parts.append(process_unit_token(part, base_units, multipliers_dict))
    return "".join(resolved_parts)

def save_mapping_to_drive(mapping_df):
    # Save updated DataFrame to a temporary file
    temp_file = "temp_mapping.xlsx"
    mapping_df.to_excel(temp_file, index=False, engine='openpyxl')
    
    # Authenticate and create a Google Drive client using PyDrive2
    gauth = GoogleAuth()
    # Attempt to load saved credentials, otherwise perform local webserver auth
    if os.path.exists("mycreds.txt"):
        gauth.LoadCredentialsFile("mycreds.txt")
    if gauth.credentials is None:
        gauth.LocalWebserverAuth()
    elif gauth.access_token_expired:
        gauth.Refresh()
    else:
        gauth.Authorize()
    gauth.SaveCredentialsFile("mycreds.txt")
    
    drive = GoogleDrive(gauth)
    # Update the file on Drive using the file ID
    file = drive.CreateFile({'id': MAPPING_FILE_ID})
    file.SetContentFile(temp_file)
    file.Upload()
    os.remove(temp_file)  # Clean up the temporary file
    return True

# ------------------ Streamlit App UI ------------------ #
st.title("Unit Processing App")

# Let user choose between two operations
operation = st.selectbox("Select Operation", options=["Get Pattern", "Add Unit"])

# Download and read the mapping file from Google Sheets
try:
    mapping_filepath = download_mapping_file()
    mapping_df, base_units, multipliers_dict = read_mapping_file(mapping_filepath)
except Exception as e:
    st.error(f"Failed to load mapping file: {e}")
    st.stop()

if operation == "Get Pattern":
    st.header("Get Pattern")
    st.write("This mode processes an input Excel file using the mapping file.")
    st.write("The mapping file is automatically loaded from Google Drive.")
    # Upload Input Excel File
    input_file = st.file_uploader("Upload Input Excel File", type=["xlsx"])
    if input_file:
        try:
            input_df = pd.read_excel(input_file)
        except Exception as e:
            st.error(f"Error reading input file: {e}")
        else:
            if "Normalized Unit" not in input_df.columns:
                st.error("Input file must contain a 'Normalized Unit' column.")
            else:
                input_df["Absolute Unit"] = input_df["Normalized Unit"].apply(
                    lambda x: resolve_compound_unit(str(x), base_units, multipliers_dict)
                )
                st.success("Processing completed!")
                towrite = BytesIO()
                input_df.to_excel(towrite, index=False, engine='openpyxl')
                towrite.seek(0)
                st.download_button(
                    label="Download Output Excel File",
                    data=towrite,
                    file_name="output.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

elif operation == "Add Unit":
    st.header("Add Unit")
    st.write("This mode lets you add a new unit to the mapping file. Only the unit symbol is required.")
    
    # Display current mapping
    st.subheader("Current Mapping File")
    st.dataframe(mapping_df)
    
    # Form to add a new unit (no multiplier required)
    with st.form(key="add_unit_form"):
        new_unit = st.text_input("Enter new Base Unit Symbol")
        submit_new = st.form_submit_button("Add New Unit")
    
    if submit_new:
        if new_unit:
            # Append new row using pd.concat instead of .append
            new_row = {"Base Unit Symbol": new_unit.strip(), "Multiplier Symbol": None}
            mapping_df = pd.concat([mapping_df, pd.DataFrame([new_row])], ignore_index=True)
            st.success("New unit added!")
            st.dataframe(mapping_df)
        else:
            st.error("The unit field is required.")
    
    # Button to download updated mapping file locally
    if st.button("Download Updated Mapping File"):
        towrite = BytesIO()
        mapping_df.to_excel(towrite, index=False, engine='openpyxl')
        towrite.seek(0)
        st.download_button(
            label="Download mapping.xlsx",
            data=towrite,
            file_name="mapping.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    
    # Button to save changes to Google Drive
    if st.button("Save Changes to Google Drive"):
        try:
            if save_mapping_to_drive(mapping_df):
                st.success("Mapping file updated on Google Drive!")
        except Exception as e:
            st.error(f"Failed to update Google Drive: {e}")
