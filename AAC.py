import streamlit as st
import pandas as pd
import re
import os
import gdown
from io import BytesIO

# ------------------ Global Variables ------------------ #
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
    # Replace with your actual mapping file ID from Google Drive
    MAPPING_FILE_ID = "YOUR_MAPPING_FILE_ID"
    mapping_url = f"https://drive.google.com/uc?export=download&id={MAPPING_FILE_ID}"
    output_path = "mapping.xlsx"
    if not os.path.exists(output_path):
        st.info("Downloading mapping file from Google Drive...")
        gdown.download(mapping_url, output_path, quiet=False)
    return output_path

def read_mapping_file(mapping_file_path):
    if not os.path.exists(mapping_file_path):
        raise FileNotFoundError(f"Error: '{mapping_file_path}' not found in the working directory.")
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

# ------------------ Streamlit App UI ------------------ #
st.title("Unit Processing App")

# Let user choose between two modes
operation = st.selectbox("Select Operation", options=["Get Pattern", "Add Unit"])

# Download the mapping file from Google Drive (if not already present)
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
    st.write("This mode lets you add a new unit to the mapping file.")
    st.write("Please ensure your input matches the mapping file headers: 'Base Unit Symbol' and 'Multiplier Symbol'.")
    
    # Display current mapping
    st.subheader("Current Mapping File")
    st.dataframe(mapping_df)
    
    # Form to add a new unit
    with st.form(key="add_unit_form"):
        new_base_unit = st.text_input("Enter new Base Unit Symbol")
        new_multiplier = st.text_input("Enter new Multiplier Symbol")
        submit_new = st.form_submit_button("Add New Unit")
    
    if submit_new:
        if new_base_unit and new_multiplier:
            # Append the new row
            new_row = {"Base Unit Symbol": new_base_unit.strip(), "Multiplier Symbol": new_multiplier.strip()}
            mapping_df = mapping_df.append(new_row, ignore_index=True)
            st.success("New unit added!")
            st.dataframe(mapping_df)
        else:
            st.error("Both fields are required.")
    
    # Provide download button for the updated mapping file
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
