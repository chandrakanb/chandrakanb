import streamlit as st
import pandas as pd
import requests
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry
import json
import re
import logging
import time
import io
from datetime import datetime
import os

# ===============================
# CONFIG
# ===============================

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s - %(levelname)s - %(message)s",
    handlers=[
        logging.FileHandler("zephyr_automation.log"),
        logging.StreamHandler()
    ]
)
logger = logging.getLogger(__name__)

BASE_URL = "https://jira.honda-ivi.com/rest/atm/1.0"
TOKEN = "NDg1ODY3NjY3ODcwOlCp9+xyU32BdsJqnmj8w1pxCy1K" 


# Token would be set as environment variable in production

HEADERS = {
    "Authorization": f"Bearer {TOKEN}",
    "content-type": "application/json"
}

AVAILABLE_FOLDERS = [
    "/PF 23/A14/AEA1415R1.260514.1/iPhone 16 Pro",
    "/PF 23/A14/AEA1415R1.260514.1/Google Pixel 9",
    "/PF 23/A14/AEA1415R1.260514.1/Samsung S25 Ultra"
]

session = requests.Session()
retries = Retry(total=3, backoff_factor=1, status_forcelist=[429, 500, 502, 503, 504])
session.mount("https://", HTTPAdapter(max_retries=retries))

RESULT_MAPPING = {
    "pass": "Pass",
    "p": "Pass",
    "fail": "Fail",
    "f": "Fail",
    "blocked": "Blocked",
    "b": "Blocked",
    "not executed": "Not Executed",
    "ne": "Not Executed",
    "na": "Not Executed",
    "": "Not Executed"
}

JIRA_USER_MAPPING = {
	"kpt_dumbre_sa" : "JIRAUSER22209",
	"kpt_sodnar_va" : "JIRAUSER22218",
	"kpt_subhash_sa" : "JIRAUSER23324",
	"kpt_chavan_ru" : "JIRAUSER23318",
	"kpt_ghadge_va" : "JIRAUSER23234",
	"kpt_bhalekar_ch" : "JIRAUSER21210",
    "kpt_rathod_na" : "JIRAUSER26036"
}

# ===============================
# API FUNCTIONS (Same as before)
# ===============================

def get_folders(project_key):
    logger.info("Loading hardcoded folders")
    folders = [{"name": path, "id": path} for path in AVAILABLE_FOLDERS]
    return folders

def create_test_cycle(project_key, folder, cycle_name, package_name, testcases):
    url = f"{BASE_URL}/testrun"
    payload = {
        "projectKey": project_key,
        "name": cycle_name,
        "description": f"Auto-created cycle for {package_name}",
        "folder": folder,
        "customFields": {"Software package name": package_name},
        "items": testcases
    }
    try:
        resp = session.post(url, headers=HEADERS, json=payload, timeout=30)
        resp.raise_for_status()
        logger.info(f"Created test cycle {cycle_name}")
        return resp.json().get("key")
    except Exception as e:
        logger.error(f"Cycle creation failed: {e}")
        st.error(f"Cycle creation failed: {e}")
        return None

def search_cycles(project_key, folder, cycle_name_base):
    url = f"{BASE_URL}/testrun/search"
    query = f'projectKey = "{project_key}" AND folder = "{folder}"'
    try:
        resp = session.get(url, headers=HEADERS, params={"query": query})
        resp.raise_for_status()
        cycles = resp.json()
        filtered_cycles = [cycle for cycle in cycles if cycle_name_base in cycle['name']]
        logger.info(f"Found {len(filtered_cycles)} cycles matching {cycle_name_base}")
        return filtered_cycles
    except Exception as e:
        logger.error(f"Failed to search cycles: {e}")
        st.error(f"Failed to search cycles: {e}")
        return []

def get_cycle_details(test_run_key):
    url = f"{BASE_URL}/testrun/{test_run_key}"
    try:
        resp = session.get(url, headers=HEADERS)
        resp.raise_for_status()
        logger.info(f"Fetched details for cycle {test_run_key}")
        return resp.json()
    except Exception as e:
        logger.error(f"Failed to fetch cycle {test_run_key}: {e}")
        st.error(f"Failed to fetch cycle {test_run_key}: {e}")
        return {}

def update_test_result(test_run_key, test_case_key, payload):
    url = f"{BASE_URL}/testrun/{test_run_key}/testcase/{test_case_key}/testresult"
    try:
        resp = session.put(url, headers=HEADERS, json=payload)
        resp.raise_for_status()
        logger.info(f"Updated result for {test_case_key} in cycle {test_run_key}")
        return True
    except Exception as e:
        logger.error(f"Error updating {test_case_key}: {e}")
        return False

def _expected_to_comment(text: str) -> str:
 
    if not isinstance(text, str):
        return ""

    # Normalize whitespace around commas and periods (optional clean-up)
    s = text.strip()

    # Replace 'should be' -> 'is' (word boundaries, case-insensitive)
    s = re.sub(r"\bshould\s+be\b", "is", s, flags=re.IGNORECASE)

    # Replace remaining standalone 'should' -> 'is' (e.g., 'playback should automatically resume')
    s = re.sub(r"\bshould\b", "is", s, flags=re.IGNORECASE)

    # Optional: compact multiple spaces created by replacements
    s = re.sub(r"\s{2,}", " ", s)

    return s
   
def parse_fail_itr_value(cell):

    if cell is None:
        return None
    s = str(cell).strip()
    if not s or "|" not in s:
        return None

    parts = [p.strip() for p in s.split("|")]
    if len(parts) < 2:
        return None

    status_token = parts[0].lower()
    if status_token not in ("fail", "f"):
        return None

    try:
        fail_index = int(parts[1])
    except Exception:
        fail_index = None

    comment = parts[2] if len(parts) >= 3 and parts[2] else None
    if comment:

        comment = re.sub(r"((?:After\s+)?step\s+\d+,?\s*)", r"\1<br>", comment, flags=re.IGNORECASE)

        comment = re.sub(r"(?<!^)\b(\d+\.\d+)\b", r"<br>\1", comment)
        
        comment = re.sub(r"(<br>\s*)+", "<br>", comment).strip()

    issue_links_raw = parts[3] if len(parts) >= 4 and parts[3] else None
    issue_links = None
    if issue_links_raw:
        issue_links = [x.strip() for x in re.split(r"[;,]", issue_links_raw) if x.strip()]

    return {
        "status": "Fail",
        "fail_index": fail_index - 1 if fail_index is not None else None,
        "comment": comment,
        "issue_links": issue_links,
    }

def build_fail_script_results_for_case(test_case_key: str, fail_index: int, fail_comment: str | None = None):
  
    url = f"{BASE_URL}/testcase/{test_case_key}"
    print("Fail Index: ",fail_index)
    print("Fail comment: ", fail_comment)
    try:
        resp = session.get(url, headers=HEADERS)
        resp.raise_for_status()
        tc = resp.json()
        steps = (tc.get("testScript") or {}).get("steps") or []
        if not isinstance(steps, list):
            steps = []

        # Collect indices and expected results
        idxs = []
        index_to_expected = {}
        for i, s in enumerate(steps):
            idx = s.get("index")
            if not isinstance(idx, int):
                idx = i  # fallback if index missing
            idxs.append(idx)

            exp = s.get("expectedResult")
            index_to_expected[idx] = exp if isinstance(exp, str) and exp.strip() else None

        # Compose scriptResults
        script_results = []
        for idx in sorted(idxs):
            if fail_index is not None and idx >= fail_index:
                entry = {"index": idx, "status": "Fail"}
                if idx == fail_index and fail_comment:
                    entry["comment"] = fail_comment
            else:
                entry = {"index": idx, "status": "Pass"}
                normalized = _expected_to_comment(index_to_expected.get(idx) or "")
                if normalized:
                    entry["comment"] = normalized
            script_results.append(entry)
        print(script_results)
        return script_results

    except requests.exceptions.RequestException as e:
        print(f" ⚠️ Could not fetch steps for {test_case_key}: {e}")
        return []


def build_pass_script_results_for_case(test_case_key):
 
    url = f"{BASE_URL}/testcase/{test_case_key}"

    try:
        resp = session.get(url, headers=HEADERS)
        resp.raise_for_status()
        tc = resp.json()

        steps = (tc.get("testScript") or {}).get("steps") or []
        if not isinstance(steps, list):
            steps = []

        # Build index -> comment mapping
        index_to_comment = {}
        for i, s in enumerate(steps):
            idx = s.get("index")
            if not isinstance(idx, int):
                idx = i  # fallback to position

            exp = s.get("expectedResult")
            comment = None
            if isinstance(exp, str) and exp.strip():
                # Normalize 'should be' / 'should' -> 'is'
                normalized = _expected_to_comment(exp)
                if normalized:
                    comment = normalized

            # Keep first non-empty comment for a given index
            if idx not in index_to_comment:
                index_to_comment[idx] = comment
            else:
                if not index_to_comment[idx] and comment:
                    index_to_comment[idx] = comment

        # Build final scriptResults sorted by index
        script_results = []
        for idx in sorted(index_to_comment.keys()):
            entry = {"index": idx, "status": "Pass"}
            if index_to_comment[idx]:
                entry["comment"] = index_to_comment[idx]
            script_results.append(entry)
        print(script_results)
        return script_results

    except requests.exceptions.RequestException as e:
        print(f"  ⚠️  Could not fetch steps for {test_case_key}: {e}")
        return []


# ===============================
# STREAMLIT UI
# ===============================

def main():
    st.set_page_config(
        page_title="Zephyr Automation Tool",
        page_icon="🚀",
        layout="wide",
        initial_sidebar_state="expanded"
    )
    
    # Custom CSS for better styling
    st.markdown("""
    <style>
    .main-header {
        font-size: 2.5rem;
        color: #1f77b4;
        text-align: center;
        margin-bottom: 2rem;
    }
    .section-header {
        font-size: 1.5rem;
        color: #2e86ab;
        border-bottom: 2px solid #2e86ab;
        padding-bottom: 0.5rem;
        margin-top: 2rem;
    }
    .success-box {
        background-color: #d4edda;
        border: 1px solid #c3e6cb;
        border-radius: 5px;
        padding: 15px;
        margin: 10px 0;
    }
    .error-box {
        background-color: #f8d7da;
        border: 1px solid #f5c6cb;
        border-radius: 5px;
        padding: 15px;
        margin: 10px 0;
    }
    .info-box {
        background-color: #d1ecf1;
        border: 1px solid #bee5eb;
        border-radius: 5px;
        padding: 15px;
        margin: 10px 0;
    }
    </style>
    """, unsafe_allow_html=True)
    
    # Header
    st.markdown('<div class="main-header">🚀 Zephyr Automation Tool</div>', unsafe_allow_html=True)
    
    # Sidebar for navigation
    st.sidebar.title("Navigation")
    app_mode = st.sidebar.radio(
        "Choose Functionality",
        ["Create Test Cycles", "Update Execution Results"]
    )
    
    # API Status Check
    st.sidebar.markdown("---")
    st.sidebar.subheader("API Status")
    if st.sidebar.button("Check Connection"):
        try:
            test_resp = session.get(f"{BASE_URL}/testrun/search", headers=HEADERS, params={"query": 'projectKey = "CNTKPIT"'})
            if test_resp.status_code == 200:
                st.sidebar.success("✅ API Connection Successful")
            else:
                st.sidebar.error("❌ API Connection Failed")
        except Exception as e:
            st.sidebar.error(f"❌ Connection Error: {e}")
    
    # Main content based on selection
    if app_mode == "Create Test Cycles":
        show_create_screen()
    else:
        show_update_screen()

def show_create_screen():
    st.markdown('<div class="section-header">Create Test Cycles</div>', unsafe_allow_html=True)
    
    if "creation_results_text" not in st.session_state:
        st.session_state.creation_results_text = None
    
    with st.form("create_cycle_form"):
        col1, col2 = st.columns(2)
        
        with col1:
            project_key = st.text_input("Project Key*", placeholder="Enter project key...")
            folder = st.selectbox("Folder*", AVAILABLE_FOLDERS)
            base_cycle_name = st.text_input("Base Cycle Name*", placeholder="e.g., RAEB1413S0.251113.1_T90A3_Short_iPhone_Cycle_2")
            
        with col2:
            package_name = st.text_input("Software Package Name*", placeholder="Enter package name...")
            iterations = st.number_input("Number of Iterations*", min_value=1, max_value=100, value=1)
            excel_file = st.file_uploader("Upload Excel File*", type=['xlsx'], help="Excel file with test cases")
        
        # Form submission
        submitted = st.form_submit_button("🚀 Create Test Cycles", use_container_width=True)
        
        if submitted:
            if not all([project_key, folder, base_cycle_name, package_name, excel_file]):
                st.error("Please fill all required fields (*)")
                return
            
            # Process the Excel file
            try:
                df = pd.read_excel(excel_file)
                
                # Validate required columns
                required_columns = ["testCaseKey", "benchEnvironment", "assignedTo", "executionMethod", "variantCode"]
                missing_columns = [col for col in required_columns if col not in df.columns]
                if missing_columns:
                    st.error(f"Missing required columns: {', '.join(missing_columns)}")
                    return
                
                if df["testCaseKey"].duplicated().any():
                    st.error("Duplicate testCaseKey found in Excel file")
                    return
                
                # Prepare test case items
                items = []
                for _, row in df.iterrows():

                    excel_user = str(row["assignedTo"]).strip().lower()

                    jira_user = JIRA_USER_MAPPING.get(excel_user)

                    if not jira_user:
                        st.error(f"User mapping not found for: {excel_user}")
                        return

                    item = {
                        "testCaseKey": row["testCaseKey"],
                        "assignedTo": jira_user,
                        "customFields": {
                            "Execution Method": row["executionMethod"],
                            "Variant Code": row["variantCode"]
                        },
                        "environment": row["benchEnvironment"]
                    }

                    items.append(item)
                    
                # Create cycles
                progress_bar = st.progress(0)
                status_text = st.empty()
                results_container = st.container()
                
                successful_creations = 0
                results = []
                
                for i in range(1, iterations + 1):
                    cycle_name = f"{base_cycle_name}_ITR_{str(i).zfill(2)}"
                    status_text.text(f"Creating {cycle_name}...")
                    
                    key = create_test_cycle(project_key, folder, cycle_name, package_name, items)
                    if key:
                        successful_creations += 1
                        results.append(f"✅ {cycle_name} → {key}")
                    else:
                        results.append(f"❌ {cycle_name} → Failed")
                    progress_bar.progress(int((i / iterations) * 100))
                    time.sleep(0.5)  # Small delay to avoid rate limiting
                
                # Display results
                status_text.text("")
                with results_container:
                    st.markdown("### 📋 Creation Results")
                    for result in results:
                        st.write(result)
                    
                    st.markdown(f"**Summary:** {successful_creations}/{iterations} cycles created successfully")
                    
                    if successful_creations > 0:
                        st.markdown('<div class="success-box">🎉 Test cycles created successfully!</div>', unsafe_allow_html=True)
                    
                    #Download results as text file
                    st.session_state.creation_results_text = "\n".join(results)
                        
            except Exception as e:
                st.error(f"Error processing Excel file: {e}")

    if st.session_state.creation_results_text:
        st.download_button(
            label="📥 Download Results Log",
            data=st.session_state.creation_results_text,
            file_name=f"cycle_creation_{datetime.now().strftime('%Y%m%d_%H%M%S')}.txt",
            mime="text/plain"
        )
    
def show_update_screen():
    st.markdown('<div class="section-header">Update Execution Results</div>', unsafe_allow_html=True)
    
    col1, col2 = st.columns([2, 1])
    
    with col1:
        with st.form("update_search_form"):
            project_key = st.text_input("Project Key*", placeholder="Enter project key...")
            folder = st.selectbox("Folder*", AVAILABLE_FOLDERS)
            build_number = st.text_input("Build Number*", placeholder="e.g., AEB1413S0.251113.1")
            
            search_submitted = st.form_submit_button("🔍 Search Test Cycles")
    
    with col2:
        excel_file = st.file_uploader("Upload Results Excel*", type=['xlsx'], 
                                    help="Excel file with ZephyrKey and ITR_xx columns")
    
    # Search results section
    if 'cycles_found' not in st.session_state:
        st.session_state.cycles_found = []
    
    if search_submitted:
        if not all([project_key, folder, build_number]):
            st.error("Please fill all required fields (*)")
        else:
            with st.spinner("Searching for test cycles..."):
                cycles = search_cycles(project_key, folder, build_number)
                st.session_state.cycles_found = cycles
                
                if cycles:
                    st.success(f"Found {len(cycles)} test cycles")
                else:
                    st.warning("No test cycles found matching your criteria")
    
    # Display found cycles
    if st.session_state.cycles_found:
        st.markdown("### 📊 Found Test Cycles")
        
        cycles_data = []
        for cycle in st.session_state.cycles_found:
            cycle_name = cycle.get("name", "")
            cycle_key = cycle.get("key", "")
            match = re.search(r'ITR_(\d+)', cycle_name, re.IGNORECASE)
            itr_column = f"ITR_{match.group(1)}" if match else "N/A"
            
            cycles_data.append({
                "Cycle Name": cycle_name,
                "Cycle Key": cycle_key,
                "Excel Column": itr_column
            })
        
        cycles_df = pd.DataFrame(cycles_data)
        st.dataframe(cycles_df, use_container_width=True)
    
    # Update results section
    if st.session_state.cycles_found and excel_file:
        st.markdown("---")
        st.markdown("### 🔄 Update Results")
        
        if st.button("🚀 Update All Results", type="primary", use_container_width=True):
            update_results(project_key, folder, build_number, excel_file, st.session_state.cycles_found)

def update_results(project_key, folder, build_number, excel_file, cycles):
    """Update test results based on Excel data"""
    
    try:
        # Read and process Excel file
        BASE_ALLOWED = [
            "ZephyrKey", "assignedTo", "executedBy", "environment",
            "Execution Method", "Variant Code"
        ]
        ITR_PREFIX = "ITR_" 
        
        df = pd.read_excel(excel_file, engine="openpyxl")
        df.columns = [c.strip() if isinstance(c, str) else c for c in df.columns]
        
        if "ZephyrKey" not in df.columns:
            st.error("Excel must contain 'ZephyrKey' column")
            return
        
        # Determine allowed ITR columns
        itr_cols = [c for c in df.columns if isinstance(c, str) and c.startswith(ITR_PREFIX)]
        allowed_now = [c for c in BASE_ALLOWED if c in df.columns] + itr_cols
        df = df.loc[:, allowed_now].copy()
        
        # Normalize data
        for c in ["assignedTo", "executedBy", "environment", "Execution Method", "Variant Code"]:
            if c in df.columns:
                df[c] = df[c].apply(lambda x: str(x).strip() if pd.notna(x) else x)
        
        # Set ZephyrKey as index
        df = df[pd.notna(df["ZephyrKey"])].copy()
        df["ZephyrKey"] = df["ZephyrKey"].astype(str).str.strip()
        df.set_index("ZephyrKey", inplace=True)
        
        if not itr_cols:
            st.error("Excel must contain ITR_xx columns (ITR_01, ITR_02, etc.)")
            return
        
    except Exception as e:
        st.error(f"Failed to read Excel file: {e}")
        return
    
    # Initialize progress tracking
    progress_bar = st.progress(0)
    status_text = st.empty()
    results_container = st.empty()
    
    summary = {"Pass": 0, "Fail": 0, "Blocked": 0, "Not Executed": 0}
    cycle_details = {}
    total_updates = 0
    all_results = []
    
    # Process each cycle
    total_cycles = len(cycles)
    
    for idx, cycle in enumerate(cycles):
        cycle_name = cycle.get("name")
        cycle_key = cycle.get("key")
        match = re.search(r'ITR_(\d+)', cycle_name)
        
        if not match:
            all_results.append(f"❌ SKIPPED: {cycle_name} - no ITR pattern found")
            continue
        
        itr_number = match.group(1)
        itr_column = f"ITR_{itr_number}"
        
        if itr_column not in df.columns:
            all_results.append(f"❌ SKIPPED: {cycle_name} - column '{itr_column}' not found in Excel")
            continue
        
        # Get cycle details
        details = get_cycle_details(cycle_key)
        cycle_details[cycle_name] = {
            "key": cycle_key,
            "test_cases": {},
            "updated": 0,
            "total": len(details.get("items", []))
        }
        
        cycle_updates = 0
        cycle_results = [f"**{cycle_name}** (Column: {itr_column})"]
        
        # Update each test case in the cycle
        for item in details.get("items", []):
            tc = item.get("testCaseKey")
            if tc in df.index:
                result_value = df.loc[tc, itr_column]
                raw_str = str(result_value).strip() if not pd.isna(result_value) else ""
                fail_info = parse_fail_itr_value(raw_str)

                if fail_info:
                    status = "Fail"
                else:
                    status = RESULT_MAPPING.get(raw_str.lower(), "Not Executed")

                payload = {"status": status}

                # Step-level results and issue links based on status
                if status == "Pass":
                    script_results = build_pass_script_results_for_case(tc)
                    if script_results:
                        payload["scriptResults"] = script_results

                elif status == "Fail":
                    # Build Fail script results if we have a step index
                    if fail_info and fail_info.get("fail_index") is not None:
                        script_results = build_fail_script_results_for_case(
                            tc,
                            fail_info["fail_index"],
                            fail_info.get("comment")
                        )
                        if script_results:
                            payload["scriptResults"] = script_results

                    # Add issue links when provided (array of issue keys)
                    if fail_info and fail_info.get("issue_links"):
                        payload["issueLinks"] = fail_info["issue_links"]

                # Add additional fields for Pass/Fail
                if status in ["Pass", "Fail"]:
                    for field in ["environment", "assignedTo", "executedBy"]:
                        if field in df.columns:
                            val = df.loc[tc, field]
                            if pd.notna(val) and str(val).strip():
                                final_value = str(val).strip()
                                if field in ["assignedTo", "executedBy"]:
                                    final_value = JIRA_USER_MAPPING.get(
                                        final_value.lower(),
                                        final_value
                                    )
                                payload[field] = final_value
            
                # Custom fields
                custom_fields_payload = {}
                for field in ["Execution Method", "Variant Code"]:
                    if field in df.columns:
                        val = df.loc[tc, field]
                        if pd.notna(val) and str(val).strip():
                            custom_fields_payload[field] = str(val).strip()
                if custom_fields_payload:
                    payload["customFields"] = custom_fields_payload

                # Update via API
                success = update_test_result(cycle_key, tc, payload)
                
                if success:
                    summary[status] += 1
                    cycle_updates += 1
                    total_updates += 1
                    cycle_details[cycle_name]["test_cases"][tc] = status
                    cycle_results.append(f"  ✅ {tc} → {status}")
                else:
                    cycle_results.append(f"  ❌ {tc} → Update failed")
            else:
                cycle_results.append(f"  ⚠️  {tc} → Not found in Excel")
        
        cycle_details[cycle_name]["updated"] = cycle_updates
        cycle_results.append(f"  **Updated:** {cycle_updates}/{cycle_details[cycle_name]['total']} test cases")
        all_results.extend(cycle_results)
        all_results.append("")  # Empty line for spacing
        
        # Update progress
        current_progress_val = (idx + 1) / total_cycles
        progress_bar.progress(current_progress_val)
        status_text.text(f"Processing {idx + 1}/{total_cycles} cycles...")
        
    # Display final results
    progress_bar.empty()
    status_text.empty()
    
    with results_container.container():
        st.markdown("### 📋 Update Results")
        
        # Summary statistics
        col1, col2, col3, col4 = st.columns(4)
        with col1:
            st.metric("Total Cycles", total_cycles)
        with col2:
            st.metric("Total Updates", total_updates)
        with col3:
            st.metric("Pass", summary["Pass"])
        with col4:
            st.metric("Fail", summary["Fail"])
        
        # Detailed results
        st.markdown("#### Detailed Log")
        for result in all_results:
            if result.startswith("**") and result.endswith("**"):
                st.markdown(result)
            elif result.startswith("  ✅"):
                st.success(result)
            elif result.startswith("  ❌"):
                st.error(result)
            elif result.startswith("  ⚠️"):
                st.warning(result)
            else:
                st.write(result)
        
        # Download results
        results_text = "\n".join(all_results)
        summary_text = f"""
Execution Summary:
Total Cycles Processed: {total_cycles}
Total Test Cases Updated: {total_updates}

Status Breakdown:
- Pass: {summary['Pass']}
- Fail: {summary['Fail']}
- Blocked: {summary['Blocked']}
- Not Executed: {summary['Not Executed']}
        """
        
        full_results = summary_text + "\n\n" + results_text
        
        st.download_button(
            label="📥 Download Full Results Log",
            data=full_results,
            file_name=f"results_update_{datetime.now().strftime('%Y%m%d_%H%M%S')}.txt",
            mime="text/plain",
            use_container_width=True
        )

if __name__ == "__main__":
    main()