import streamlit as st
import pandas as pd
import re
import os
import gdown
import json
from io import BytesIO
from pydrive2.auth import GoogleAuth
from pydrive2.drive import GoogleDrive

# ------------------ Global Constants ------------------ #
MAPPING_FILE_ID = "1QP1XnxyDEgfxYfgBg_mf2ngXNfm9O8s5"

MULTIPLIER_MAPPING = {
    'k': 1e3,
    'M': 1e6,
    'G': 1e9,
    'T': 1e12,
    'P': 1e15,
    'E': 1e18,
    'Z': 1e21,
    'Y': 1e24,
    'd': 1e-1,
    'c': 1e-2,
    'm': 1e-3,
    'µ': 1e-6,
    'n': 1e-9,
    'p': 1e-12,
    'f': 1e-15,
    'a': 1e-18,
    'z': 1e-21,
    'y': 1e-24
}

# ------------------ Helper Functions ------------------ #
def download_mapping_file():
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
    # Save updated mapping to a temporary file.
    temp_file = "temp_mapping.xlsx"
    mapping_df.to_excel(temp_file, index=False, engine='openpyxl')
    
    # Initialize GoogleAuth without a local settings file.
    gauth = GoogleAuth(settings_file=None)
    
    # Load client configuration from st.secrets.
    try:
        client_config_full = json.loads(st.secrets["google"]["client_secrets"])
        st.write("DEBUG: Full client_config from secrets:", client_config_full)
    except Exception as e:
        st.error("DEBUG: Error loading client_secrets from st.secrets: " + str(e))
        raise

    # Check if the "installed" key exists; if so, use it.
    if "installed" in client_config_full:
        client_config = client_config_full["installed"]
        st.write("DEBUG: Using client_config['installed']:", client_config)
    else:
        client_config = client_config_full
        st.write("DEBUG: Using full client_config:", client_config)
    
    # Set client configuration settings for PyDrive2.
    gauth.settings["client_config_backend"] = "settings"
    gauth.settings["client_config"] = client_config
    
    # Debug: Check for required keys.
    required_keys = ["client_id", "client_secret", "project_id", "auth_uri", "token_uri", "auth_provider_x509_cert_url", "redirect_uris"]
    missing_keys = [key for key in required_keys if key not in gauth.settings["client_config"]]
    if missing_keys:
        st.error("DEBUG: Missing keys in client config: " + ", ".join(missing_keys))
        raise Exception("Insufficient client config in settings: missing " + ", ".join(missing_keys))
    
    # Load saved credentials if available; otherwise perform authentication.
    if os.path.exists("mycreds.txt"):
        gauth.LoadCredentialsFile("mycreds.txt")
        st.write("DEBUG: Loaded saved credentials from mycreds.txt")
    if gauth.credentials is None:
        st.write("DEBUG: No credentials found; performing LocalWebserverAuth")
        gauth.LocalWebserverAuth()  # Opens a browser window for authentication.
    elif gauth.access_token_expired:
        st.write("DEBUG: Credentials expired; refreshing")
        gauth.Refresh()
    else:
        st.write("DEBUG: Credentials valid; authorizing")
        gauth.Authorize()
    gauth.SaveCredentialsFile("mycreds.txt")
    
    drive = GoogleDrive(gauth)
    st.write("DEBUG: Uploading temporary file:", temp_file)
    file = drive.CreateFile({'id': MAPPING_FILE_ID})
    file.SetContentFile(temp_file)
    file.Upload()
    st.write("DEBUG: File uploaded to Google Drive.")
    os.remove(temp_file)
    return True

# ------------------ Streamlit App UI ------------------ #
st.title("Unit Processing App")

operation = st.selectbox("Select Operation", options=["Get Pattern", "Add Unit"])

try:
    mapping_filepath = download_mapping_file()
    mapping_df, base_units, multipliers_dict = read_mapping_file(mapping_filepath)
except Exception as e:
    st.error(f"Failed to load mapping file: {e}")
    st.stop()

if operation == "Get Pattern":
    st.header("Get Pattern")
    st.write("This mode processes an input Excel file using the mapping file (loaded from Google Drive).")
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
    st.subheader("Current Mapping File")
    st.dataframe(mapping_df)
    
    with st.form(key="add_unit_form"):
        new_unit = st.text_input("Enter new Base Unit Symbol")
        submit_new = st.form_submit_button("Add New Unit")
    
    if submit_new:
        if new_unit:
            new_row = {"Base Unit Symbol": new_unit.strip(), "Multiplier Symbol": None}
            mapping_df = pd.concat([mapping_df, pd.DataFrame([new_row])], ignore_index=True)
            st.success("New unit added!")
            st.dataframe(mapping_df)
        else:
            st.error("The unit field is required.")
    
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
    
    if st.button("Save Changes to Google Drive"):
        try:
            if save_mapping_to_drive(mapping_df):
                st.success("Mapping file updated on Google Drive!")
        except Exception as e:
            st.error(f"Failed to update Google Drive: {e}")
