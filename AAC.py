import streamlit as st
import pandas as pd
import re
import os
import json
import base64
import requests
from io import BytesIO

############################
#  MULTIPLIER DICTIONARY   #
############################

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

############################
#  GITHUB HELPER FUNCTIONS #
############################

def download_mapping_file_from_github() -> pd.DataFrame:
    """
    Downloads 'mapping.xlsx' from a GitHub repo specified in secrets,
    returns a DataFrame parsed from that file.
    """
    st.write("DEBUG: Downloading mapping.xlsx from GitHub...")
    github_token = st.secrets["github"]["token"]
    owner = st.secrets["github"]["owner"]
    repo = st.secrets["github"]["repo"]
    file_path = st.secrets["github"]["file_path"]

    url = f"https://api.github.com/repos/{owner}/{repo}/contents/{file_path}"
    headers = {
        "Authorization": f"token {github_token}",
        "Accept": "application/vnd.github.v3+json"
    }

    response = requests.get(url, headers=headers)
    if response.status_code == 200:
        content_json = response.json()
        encoded_content = content_json["content"]
        decoded_bytes = base64.b64decode(encoded_content)

        local_file = "mapping.xlsx"
        with open(local_file, "wb") as f:
            f.write(decoded_bytes)

        # Now parse the local file into a DataFrame
        try:
            df = pd.read_excel(local_file)
        except Exception as e:
            st.error(f"Failed to parse downloaded mapping file: {e}")
            st.stop()
        os.remove(local_file)  # clean up local file
        st.write("DEBUG: Download successful. mapping_df shape:", df.shape)
        return df
    else:
        st.error(f"Failed to download file from GitHub: {response.status_code} {response.text}")
        st.stop()

def update_mapping_file_on_github(mapping_df: pd.DataFrame) -> bool:
    """
    Updates 'mapping.xlsx' on GitHub using a PUT request to the GitHub API.
    """
    st.write("DEBUG: Attempting to update mapping.xlsx on GitHub.")
    st.write("DEBUG: DataFrame shape before upload:", mapping_df.shape)

    github_token = st.secrets["github"]["token"]
    owner = st.secrets["github"]["owner"]
    repo = st.secrets["github"]["repo"]
    file_path = st.secrets["github"]["file_path"]

    # 1) Save DF to local file
    temp_file = "mapping.xlsx"
    mapping_df.to_excel(temp_file, index=False, engine='openpyxl')

    # 2) Encode local file in base64
    with open(temp_file, "rb") as f:
        content_bytes = f.read()
    encoded_content = base64.b64encode(content_bytes).decode("utf-8")

    # 3) Get the current file's SHA
    url = f"https://api.github.com/repos/{owner}/{repo}/contents/{file_path}"
    headers = {
        "Authorization": f"token {github_token}",
        "Accept": "application/vnd.github.v3+json"
    }
    current_response = requests.get(url, headers=headers)
    sha = None
    if current_response.status_code == 200:
        sha = current_response.json().get("sha")
        st.write("DEBUG: Current file SHA:", sha)
    else:
        st.write("DEBUG: No existing file found. Creating a new one...")

    # 4) Prepare data payload
    data = {
        "message": "Update mapping file via Streamlit app",
        "content": encoded_content
    }
    if sha:
        data["sha"] = sha

    # 5) PUT request to update file
    update_response = requests.put(url, headers=headers, json=data)
    os.remove(temp_file)  # clean up local file

    if update_response.status_code in [200, 201]:
        st.write("DEBUG: Update/creation successful:", update_response.status_code)
        return True
    else:
        st.error(f"Failed to update file on GitHub: {update_response.status_code} {update_response.text}")
        return False

############################
#   UNIT-PROCESSING LOGIC  #
############################

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
        after = token[1:]
        stripped = after.strip()
        if stripped == "":
            return "$"
        if stripped in base_units:
            return "$" + after
        for prefix in sorted(multipliers_dict.keys(), key=len, reverse=True):
            if stripped.startswith(prefix):
                possible = stripped[len(prefix):]
                if possible in base_units:
                    idx = after.find(prefix)
                    if idx != -1:
                        if idx == 1 and after[0] == " ":
                            preserved = after[:0] + after[idx + len(prefix):]
                        else:
                            preserved = after[:idx] + after[idx + len(prefix):]
                        return "$" + preserved
        return f"Error: Undefined unit '{stripped}' (no recognized prefix)"
    else:
        stripped = token.strip()
        if stripped in base_units:
            return "$" + stripped
        for prefix in sorted(multipliers_dict.keys(), key=len, reverse=True):
            if stripped.startswith(prefix):
                possible = stripped[len(prefix):]
                if possible in base_units:
                    idx = token.find(prefix)
                    if idx != -1:
                        if idx > 0 and token[idx-1] == " ":
                            if idx == 1:
                                preserved = token[:0] + token[idx + len(prefix):]
                            else:
                                preserved = token[:idx] + token[idx + len(prefix):]
                        else:
                            preserved = token[:idx] + token[idx + len(prefix):]
                        return "$" + preserved
        return f"Error: Undefined unit '{stripped}' (no recognized prefix)"

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
    lead = m.group('lead')
    numeric = m.group('numeric')
    space1 = m.group('space1')
    unit_part = m.group('unit')
    space2 = m.group('space2')
    paren = m.group('paren') if m.group('paren') else ""
    trail = m.group('trail')
    core = unit_part.strip()
    left_ws = re.match(r'^\s*', unit_part).group(0) or ""
    right_ws = re.search(r'\s*$', unit_part).group(0) or ""
    processed = process_unit_token_no_paren(core, base_units, multipliers_dict)
    new_unit = left_ws + processed + right_ws
    if "ohm" in core.lower():
        if new_unit.startswith("$") and not new_unit.startswith("$ "):
            new_unit = "$ " + new_unit[1:].lstrip()
        if numeric and not space1:
            space1 = " "
    return f"{lead}{numeric}{space1}{new_unit}{space2}{paren}{trail}"

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

############################
#   MAIN STREAMLIT APP     #
############################

st.title("Unit Processing App (GitHub-based)")

# Use session_state to store the DataFrame so we don't re-download after each interaction.
if "mapping_df" not in st.session_state:
    st.session_state["mapping_df"] = download_mapping_file_from_github()

mapping_df = st.session_state["mapping_df"]

# Validate columns quickly
required_cols = {"Base Unit Symbol", "Multiplier Symbol"}
if not required_cols.issubset(mapping_df.columns):
    st.error(f"Mapping file must contain columns: {required_cols}")
    st.stop()

base_units = {str(u).strip() for u in mapping_df["Base Unit Symbol"].dropna().unique()}

operation = st.selectbox("Select Operation", ["Get Pattern", "Manage Units"])

if operation == "Get Pattern":
    st.header("Get Pattern")
    st.write("This mode processes an input Excel file using the mapping file (loaded from GitHub).")
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
                    lambda x: resolve_compound_unit(str(x), base_units, MULTIPLIER_MAPPING)
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

elif operation == "Manage Units":
    st.header("Manage Units")
    st.write("This mode lets you add or remove units from the mapping file. Only the unit symbol is required for adding.")

    st.subheader("Current Mapping File")
    st.dataframe(st.session_state["mapping_df"])

    # --- Add a new unit ---
    with st.form("add_unit_form"):
        new_unit = st.text_input("Enter new Base Unit Symbol")
        submit_new = st.form_submit_button("Add New Unit")

    if submit_new:
        if new_unit.strip():
            new_row = {"Base Unit Symbol": new_unit.strip(), "Multiplier Symbol": None}
            st.session_state["mapping_df"] = pd.concat(
                [st.session_state["mapping_df"], pd.DataFrame([new_row])],
                ignore_index=True
            )
            st.success(f"New unit '{new_unit.strip()}' added!")
            st.dataframe(st.session_state["mapping_df"])
        else:
            st.error("The unit field is required.")

    # --- Delete a unit ---
    existing_units = st.session_state["mapping_df"]["Base Unit Symbol"].dropna().unique().tolist()
    if existing_units:
        to_delete = st.selectbox("Select a unit to delete", ["--Select--"] + existing_units)
        if st.button("Delete Selected Unit"):
            if to_delete == "--Select--":
                st.warning("Please select a valid unit to delete.")
            else:
                before_shape = st.session_state["mapping_df"].shape
                st.session_state["mapping_df"] = st.session_state["mapping_df"][
                    st.session_state["mapping_df"]["Base Unit Symbol"] != to_delete
                ]
                after_shape = st.session_state["mapping_df"].shape
                st.success(f"Unit '{to_delete}' has been deleted. (Rows before: {before_shape}, after: {after_shape})")
                st.dataframe(st.session_state["mapping_df"])
    else:
        st.info("No units available to delete.")

    # Option to download updated file locally
    if st.button("Download Updated Mapping File"):
        towrite = BytesIO()
        st.session_state["mapping_df"].to_excel(towrite, index=False, engine='openpyxl')
        towrite.seek(0)
        st.download_button(
            label="Download mapping.xlsx",
            data=towrite,
            file_name="mapping.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    # Save changes back to GitHub
    if st.button("Save Changes to GitHub"):
        st.write("DEBUG: Attempting to save changes to GitHub. DF shape:", st.session_state["mapping_df"].shape)
        if update_mapping_file_on_github(st.session_state["mapping_df"]):
            st.success("Mapping file updated on GitHub!")
        else:
            st.error("Failed to update mapping file on GitHub.")
