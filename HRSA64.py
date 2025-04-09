import streamlit as st
import re
import pandas as pd
from millify import millify # shortens values (10_000 ---> 10k)
from streamlit_extras.metric_cards import style_metric_cards # beautify metric card with css
from datetime import datetime, timedelta
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import json

scope = ["https://spreadsheets.google.com/feeds", 'https://www.googleapis.com/auth/spreadsheets', "https://www.googleapis.com/auth/drive.file", "https://www.googleapis.com/auth/drive"]
#creds = ServiceAccountCredentials.from_json_keyfile_name('client_secret.json', scope)
# Use Streamlit's secrets management
creds_dict = st.secrets["gcp_service_account"]
# Extract individual attributes needed for ServiceAccountCredentials
credentials = {
    "type": creds_dict.type,
    "project_id": creds_dict.project_id,
    "private_key_id": creds_dict.private_key_id,
    "private_key": creds_dict.private_key,
    "client_email": creds_dict.client_email,
    "client_id": creds_dict.client_id,
    "auth_uri": creds_dict.auth_uri,
    "token_uri": creds_dict.token_uri,
    "auth_provider_x509_cert_url": creds_dict.auth_provider_x509_cert_url,
    "client_x509_cert_url": creds_dict.client_x509_cert_url,
}

# Create JSON string for credentials
creds_json = json.dumps(credentials)

# Load credentials and authorize gspread
creds = ServiceAccountCredentials.from_json_keyfile_dict(json.loads(creds_json), scope)
client = gspread.authorize(creds)

# Example usage: Fetch data from Google Sheets
try:
    spreadsheet1 = client.open('Example_TA_Request')
    worksheet1 = spreadsheet1.worksheet('Main')
    df = pd.DataFrame(worksheet1.get_all_records())
except Exception as e:
    st.error(f"Error fetching data from Google Sheets: {str(e)}")

df['Submit Date'] = pd.to_datetime(df['Submit Date'], errors='coerce')
df["Phone Number"] = df["Phone Number"].astype(str)
def format_phone(phone_str):
    # Remove non-digit characters
    digits = re.sub(r'\D', '', phone_str)
    
    # Format if it's 10 digits long
    if len(digits) == 10:
        return f"({digits[:3]}) {digits[3:6]}-{digits[6:]}"
    elif len(digits) == 11 and digits.startswith("1"):
        return f"+1 ({digits[1:4]}) {digits[4:7]}-{digits[7:]}"
    else:
        return phone_str  # Return original if not standard length

# Apply formatting
df["Phone Number"] = df["Phone Number"].astype(str).apply(format_phone)

# --- Demo user database
USERS = {
    "jw2104@georgetown.com": {"password": "qin88251216", "role": "Coordinator", "name":"Jiaqin"},
    "jiaqinwu@georgetown.com": {"password": "qin88251216", "role": "Assignee/Staff", "name":"MM"}
}
lis_location = ["Maricopa Co. - Arizona", "Alameda Co. - California", "Los Angeles Co. - California", "Orange Co. - California", "Riverside Co. - California",\
                "Sacramento Co. - California", "San Bernadino Co. -California", "San Diego Co. - California", "San Francisco Co. - California",\
                "Broward Co. - Florida", "Duval Co. - Florida", "Hillsborough Co. - Florida", "Miami-Dade Co. - Florida","Orange Co. - Florida",\
                "Palm Beach Co. - Florida", "Pinellas Co. - Florida", "Cobb Co. - Georgia", "Dekalb Co. - Georgia", "Fulton Co. - Georgia",\
                "Gwinnett Co. - Georgia", "Cook Co. - Illinois", "Marion Co. - Indiana", "East Baton Rough Parish - Louisiana",\
                "Orleans Parish - Louisiana", "Baltimore City - Maryland", "Montgomery Co. - Maryland", "Prince George's Co. - Maryland",\
                "Suffolk Co. - Massachusetts", "Wayne Co. - Michigan", "Clark Co. - Neveda", "Essex Co. - New Jersey","Hudson Co. - New Jersey",\
                "Bronx Co. - New York", "Kings Co. - New York", "New York Co. - New York", "Queens Co. - New York", "Mecklenburg Co. - North Carolina",\
                "Cuyahoga Co. - Ohio", "Franklin Co. - Ohio", "Hamilton Co. - Ohio", "Philadelphia Co. - Pennsylvania", "Shelby Co. - Tennessee",\
                "Bexar Co. - Texas", "Dallas Co. - Texas","Harris Co. - Texas", "Tarrant Co. - Texas","Travis Co. - Texas","King Co. - Washington",\
                "Washington, DC", "San Juan Municipio - Puerto Rico", "Alabama", "Arkansas","Kentucky","Mississippi","Missouri","Oklahoma","South Carolina"]

# --- Initialize session state
if "authenticated" not in st.session_state:
    st.session_state.authenticated = False
if "role" not in st.session_state:
    st.session_state.role = None
if "user_email" not in st.session_state:
    st.session_state.user_email = ""

# --- Role selection
if st.session_state.role is None:
    st.title("Welcome to the TA Request System")

    role = st.selectbox(
        "Select your role",
        ["Requester", "Coordinator", "Assignee/Staff"],
        index=None,
        placeholder="Select option..."
    )

    if role:
        st.session_state.role = role
        st.rerun()

# --- Show view based on role
else:
    st.sidebar.title(f"Role: {st.session_state.role}")
    st.sidebar.button("🔄 Switch Role", on_click=lambda: st.session_state.update({
        "authenticated": False,
        "role": None,
        "user_email": ""
    }))
    # --- Submit button styling (CSS injection)
    st.markdown("""
        <style>
        .stButton > button {
            width: 100%;
            background-color: #cdb4db;
            color: black;
            font-weight: 600;
            border-radius: 8px;
            padding: 0.6em;
            margin-top: 1em;
        }
        </style>
    """, unsafe_allow_html=True)

    # --- Requester: No login needed
    if st.session_state.role == "Requester":
        st.header("📥 NASTAD Technical Assistance Form")
        st.write("Please complete this form to request Technical Assistance from NASTAD's Health Care Access and Health Systems Integration teams. We will review your request and will be in touch within 1-2 business days. You will receive an email from a TA Coordinator to schedule a time to gather more details about your needs. Once we have this information, we will assign a TA Lead to support you.")
        # Add requester form here
        col1, col2 = st.columns(2)
        with col1:
            name = st.text_input("Name *",placeholder="Enter text")
        with col2:
            title = st.text_input("Title/Position *",placeholder='Enter text')
        col3, col4 = st.columns(2)
        with col3:
            organization = st.selectbox(
                "Organization *",
                ["CGHPI", "HRSA", "NASTAD"],
                index=None,
                placeholder="Select option..."
            )
        with col4:
            location = st.selectbox(
                "Location *",
                lis_location,
                index=None,
                placeholder="Select option..."
            )
        col5, col6 = st.columns(2)
        with col5:
            email = st.text_input("Email Address *",placeholder="Enter email")
        with col6:
            phone = st.text_input("Phone Number *",placeholder="(201) 555-0123")    

        col7, col8 = st.columns(2)
        with col7:
            focus_area = st.selectbox(
                "TA Focus Area *",
                ["Housing", "Prevention", "Substance Abuse","Rapid Start","Telehealth/Telemedicine","Data Sharing"],
                index=None,
                placeholder="Select option..."
            )
        with col8:
            type_TA = st.selectbox(
                "What Style of TA is needed *",
                ["In-Person","Virtual","Hybrid (Combination of in-person and virtual)"],
                index=None,
                placeholder="Select option..."
            )
        col9, col10 = st.columns(2)
        with col9:
            due_date = st.date_input(
                "Target Due Date *",
                value=None
            )

        ta_description = st.text_area("TA Description *", placeholder='Enter text', height=150) 
        priority_status = st.selectbox(
                "Priority Status *",
                ["Critical","High","Normal","Low"],
                index=None,
                placeholder="Select option..."
            )
        
        # --- Submit button styling (CSS injection)
        st.markdown("""
            <style>
            .stButton > button {
                width: 100%;
                background-color: #cdb4db;
                color: black;
                font-weight: 600;
                border-radius: 8px;
                padding: 0.6em;
                margin-top: 1em;
            }
            </style>
        """, unsafe_allow_html=True)

        # --- Submit logic
        if st.button("Submit"):
            email_pattern = r'^[\w\.-]+@[\w\.-]+\.\w+$'
            phone_pattern = r'^\+?[\d\s\-\(\)]{7,20}$'

            errors = []

            # Required field checks
            if not name: errors.append("Name is required.")
            if not title: errors.append("Title/Position is required.")
            if not organization: errors.append("Organization must be selected.")
            if not location: errors.append("Location must be selected.")
            if not email or not re.match(email_pattern, email):
                errors.append("Please enter a valid email address.")
            if not phone or not re.match(phone_pattern, phone):
                errors.append("Please enter a valid phone number.")
            if not focus_area: errors.append("TA Focus Area must be selected.")
            if not type_TA: errors.append("TA Style must be selected.")
            if not due_date: errors.append("Target Due Date is required.")
            if not ta_description: errors.append("TA Description is required.")
            if not priority_status: errors.append("Priority Status must be selected.")

            # Show warnings or success
            if errors:
                for error in errors:
                    st.warning(error)
            else:
                # Add logic here to save to Google Sheet or database (add two more column "Submit Date" and "Status" marked as "Submitted" too)
                # Prepare the new row
                new_row = {
                    'Jurisdiction': location,
                    'Organization': organization,
                    'Name': name,
                    'Title/Position': title,
                    'Email Address': email,
                    "Phone Number": phone,
                    "TA Type": type_TA,
                    "Targeted Due Date": due_date.strftime("%Y-%m-%d"),
                    "TA Description":ta_description,
                    "Priority": priority_status,
                    "Submit Date": datetime.today().strftime("%Y-%m-%d"),
                    "Status": "Submitted"
                }
                new_data = pd.DataFrame([new_row])

                try:
                    # Append new data to Google Sheet
                    updated_sheet = pd.concat([df, new_data], ignore_index=True)
                    worksheet1.update([updated_sheet.columns.values.tolist()] + updated_sheet.values.tolist())
                    st.success("✅ Submission successful!")
                except Exception as e:
                    st.error(f"Error updating Google Sheets: {str(e)}")
                
                


    # --- Coordinator or Staff: Require login
    elif st.session_state.role in ["Coordinator", "Assignee/Staff"]:
        if not st.session_state.authenticated:
            st.subheader("🔐 Login Required")

            email = st.text_input("Email")
            password = st.text_input("Password", type="password")
            login = st.button("Login")

            if login:
                user = USERS.get(email)
                if user and user["password"] == password and user["role"] == st.session_state.role:
                    st.session_state.authenticated = True
                    st.session_state.user_email = email
                    st.success("Login successful!")
                    st.rerun()
                else:
                    st.error("Invalid credentials or role mismatch.")

        else:
            if st.session_state.role == "Coordinator":
                user_info = USERS.get(st.session_state.user_email)
                st.header("📬 Coordinator Dashboard")
                # Personalized greeting
                if user_info and "name" in user_info:
                    st.markdown(f"#### 👋 Welcome, {user_info['name']}!")
                col1, col2, col3 = st.columns(3)
                # create column span
                today = datetime.today()
                last_week = today - timedelta(days=7)
                last_month = today - timedelta(days=30)
                undone_request = df[df['Status'] == 'Submitted'].shape[0]
                pastweek_request = df[df['Submit Date'] >= last_week].shape[0]
                pastmonth_request = df[df['Submit Date'] >= last_month].shape[0]
                col1.metric(label="# of Unassigned Requests", value= millify(undone_request, precision=2))
                col2.metric(label="# of Requests from past week", value= millify(pastweek_request, precision=2))
                col3.metric(label="# of Requests from past month", value= millify(pastmonth_request, precision=2))
                style_metric_cards(border_left_color="#DBF227")

                staff_list = ["MM", "KK", "LL"]

                # Filter submitted requests
                submitted_requests = df[df["Status"] == "Submitted"].copy()

                st.subheader("📋 Unassigned Requests")

                if submitted_requests.empty:
                    st.info("No submitted requests at the moment.")
                else:
                    # Define custom priority order
                    priority_order = {"Critical": 1, "High": 2, "Normal": 3, "Low": 4}

                    # Create a temporary column for sort priority
                    submitted_requests["PriorityOrder"] = submitted_requests["Priority"].map(priority_order)

                    # Convert date columns if needed
                    submitted_requests["Submit Date"] = pd.to_datetime(submitted_requests["Submit Date"], errors='coerce')
                    submitted_requests["Targeted Due Date"] = pd.to_datetime(submitted_requests["Targeted Due Date"], errors='coerce')

                    # Format dates to "YYYY-MM-DD" for display
                    submitted_requests["Submit Date"] = submitted_requests["Submit Date"].dt.strftime("%Y-%m-%d")
                    submitted_requests["Targeted Due Date"] = submitted_requests["Targeted Due Date"].dt.strftime("%Y-%m-%d")

                    # Sort by custom priority, then submit date, then due date
                    submitted_requests_sorted = submitted_requests.sort_values(
                        by=["PriorityOrder", "Submit Date", "Targeted Due Date"],
                        ascending=[True, True, True]
                    )

                    # Display clean table (exclude PriorityOrder column)
                    st.dataframe(submitted_requests_sorted[[
                        "Jurisdiction", "Organization", "Name", "Title/Position", "Email Address", "Phone Number",
                        "Focus Area", "TA Type", "Submit Date", "Targeted Due Date", "Priority", "TA Description"
                    ]].reset_index(drop=True))

                    # Select request by index (row number in submitted_requests)
                    request_indices = submitted_requests.index.tolist()
                    selected_request_index = st.selectbox(
                        "Select a request to assign",
                        options=request_indices,
                        format_func=lambda idx: f"{submitted_requests.at[idx, 'Name']} | {submitted_requests.at[idx, 'Jurisdiction']}",
                    )

                    # Select coach
                    selected_coach = st.selectbox(
                        "Assign a coach",
                        options=staff_list,
                        index=None,
                        placeholder="Select option..."
                    )
                    

                    # Assign button
                    if st.button("✅ Assign Coach and Start TA"):
                        try:
                            # Create a copy to avoid modifying the original df directly (optional but safe)
                            updated_df = df.copy()

                            # Update the selected row
                            updated_df.loc[selected_request_index, "Assigned Coach"] = selected_coach
                            updated_df.loc[selected_request_index, "Status"] = "In Progress"
                            updated_df.loc[selected_request_index, "Assigned Date"] = datetime.today().strftime("%Y-%m-%d")

                            # Push full updated DataFrame to Google Sheets
                            worksheet1.update([updated_df.columns.values.tolist()] + updated_df.values.tolist())

                            st.success(f"Coach {selected_coach} assigned! Status updated to 'In Progress'.")
                            st.rerun()

                        except Exception as e:
                            st.error(f"Error updating Google Sheets: {str(e)}")

                    # --- Submit button styling (CSS injection)
                    st.markdown("""
                        <style>
                        .stButton > button {
                            width: 100%;
                            background-color: #cdb4db;
                            color: black;
                            font-weight: 600;
                            border-radius: 8px;
                            padding: 0.6em;
                            margin-top: 1em;
                        }
                        </style>
                    """, unsafe_allow_html=True)

                    st.subheader("🚧 In-progress Requests")

                    # Filter "In Progress" requests
                    in_progress_df = df[df["Status"] == "In Progress"].copy()

                    if in_progress_df.empty:
                        st.info("No requests currently in progress.")
                    else:
                        # Convert date columns
                        in_progress_df["Assigned Date"] = pd.to_datetime(in_progress_df["Assigned Date"], errors="coerce")
                        in_progress_df["Targeted Due Date"] = pd.to_datetime(in_progress_df["Targeted Due Date"], errors="coerce")
                        in_progress_df['Expected Duration (Days)'] = (in_progress_df["Targeted Due Date"]-in_progress_df["Assigned Date"]).dt.days

                        # Format dates
                        in_progress_df["Assigned Date"] = in_progress_df["Assigned Date"].dt.strftime("%Y-%m-%d")
                        in_progress_df["Targeted Due Date"] = in_progress_df["Targeted Due Date"].dt.strftime("%Y-%m-%d")
                        

                        # --- Filters
                        st.markdown("#### 🔍 Filter Options")

                        col1, col2, col3 = st.columns(3)
                        with col1:
                            priority_filter = st.multiselect(
                                "Filter by Priority",
                                options=in_progress_df["Priority"].unique(),
                                default=in_progress_df["Priority"].unique(), key='in1'
                            )

                        with col2:
                            ta_type_filter = st.multiselect(
                                "Filter by TA Type",
                                options=in_progress_df["TA Type"].unique(),
                                default=in_progress_df["TA Type"].unique(), key='in2'
                            )

                        with col3:
                            focus_area_filter = st.multiselect(
                                "Filter by Focus Area",
                                options=in_progress_df["Focus Area"].unique(),
                                default=in_progress_df["Focus Area"].unique(), key='in3'
                            )

                        # Apply filters
                        filtered_df = in_progress_df[
                            (in_progress_df["Priority"].isin(priority_filter)) &
                            (in_progress_df["TA Type"].isin(ta_type_filter)) &
                            (in_progress_df["Focus Area"].isin(focus_area_filter))
                        ]



                        # Display filtered table
                        st.dataframe(filtered_df[[
                            "Jurisdiction", "Organization", "Name", "Title/Position", "Email Address", "Phone Number",
                            "Focus Area", "TA Type", "Assigned Date", "Targeted Due Date","Expected Duration (Days)","Priority", "Assigned Coach", "TA Description"
                        ]].sort_values(by="Expected Duration (Days)").reset_index(drop=True))

                        # Select request by index (row number in submitted_requests)
                        request_indices1 = filtered_df.index.tolist()
                        selected_request_index1 = st.selectbox(
                            "Select a request to comment",
                            options=request_indices1,
                            format_func=lambda idx: f"{filtered_df.at[idx, 'Name']} | {filtered_df.at[idx, 'Jurisdiction']}",
                        )

                        # Input + submit comment
                        comment_input = st.text_area("Comments", placeholder="Enter comments", height=150, key='comm')

                        if st.button("✅ Submit Comments"):
                            try:
                                # Find the actual row index in the original df (map back using ID or index)
                                selected_row_global_index = filtered_df.loc[selected_request_index1].name

                                # Copy and update df
                                updated_df = df.copy()
                                updated_df.loc[selected_row_global_index, "Coordinator Comment"] = comment_input

                                # Push to Google Sheets
                                worksheet1.update([updated_df.columns.values.tolist()] + updated_df.values.tolist())

                                st.success("💬 Comment saved and synced with Google Sheets.")
                                st.rerun()

                            except Exception as e:
                                st.error(f"Error updating Google Sheets: {str(e)}")


                    st.subheader("✅ Completed Requests")

                    # Filter "Completed" requests
                    complete_df = df[df["Status"] == "Completed"].copy()

                    if complete_df.empty:
                        st.info("No requests currently completed.")
                    else:
                        # Convert date columns
                        complete_df["Assigned Date"] = pd.to_datetime(complete_df["Assigned Date"], errors="coerce")
                        complete_df["Close Date"] = pd.to_datetime(complete_df["Close Date"], errors="coerce")
                        complete_df["Targeted Due Date"] = pd.to_datetime(complete_df["Targeted Due Date"], errors="coerce")
                        complete_df['Expected Duration (Days)'] = (complete_df["Targeted Due Date"]-complete_df["Assigned Date"]).dt.days
                        complete_df['Actual Duration (Days)'] = (complete_df["Close Date"]-complete_df["Assigned Date"]).dt.days

                        # Format dates
                        complete_df["Assigned Date"] = complete_df["Assigned Date"].dt.strftime("%Y-%m-%d")
                        complete_df["Targeted Due Date"] = complete_df["Targeted Due Date"].dt.strftime("%Y-%m-%d")
                        complete_df["Close Date"] = complete_df["Close Date"].dt.strftime("%Y-%m-%d")

                        # --- Filters
                        st.markdown("#### 🔍 Filter Options")

                        col1, col2, col3 = st.columns(3)
                        with col1:
                            priority_filter1 = st.multiselect(
                                "Filter by Priority",
                                options=complete_df["Priority"].unique(),
                                default=complete_df["Priority"].unique(), key='com1'
                            )

                        with col2:
                            ta_type_filter1 = st.multiselect(
                                "Filter by TA Type",
                                options=complete_df["TA Type"].unique(),
                                default=complete_df["TA Type"].unique(), key='com2'
                            )

                        with col3:
                            focus_area_filter1 = st.multiselect(
                                "Filter by Focus Area",
                                options=complete_df["Focus Area"].unique(),
                                default=complete_df["Focus Area"].unique(), key='com3'
                            )

                        # Apply filters
                        filtered_df1 = complete_df[
                            (complete_df["Priority"].isin(priority_filter1)) &
                            (complete_df["TA Type"].isin(ta_type_filter1)) &
                            (complete_df["Focus Area"].isin(focus_area_filter1))
                        ]

                        # Display filtered table
                        st.dataframe(filtered_df1[[
                            "Jurisdiction", "Organization", "Name", "Title/Position", "Email Address", "Phone Number",
                            "Focus Area", "TA Type", "Assigned Date", "Targeted Due Date", "Close Date", "Expected Duration (Days)",
                            'Actual Duration (Days)', "Priority", "Assigned Coach", "TA Description"
                        ]].reset_index(drop=True))

            elif st.session_state.role == "Assignee/Staff":
                # Add staff content here
                staff_name = USERS[st.session_state.user_email]["name"]

                st.markdown(f"#### 👋 Welcome, {staff_name}!")

                st.header("👷 Staff Dashboard")

                # Filter requests assigned to current staff and In Progress
                staff_df = df[(df["Assigned Coach"] == staff_name) & (df["Status"] == "In Progress")].copy()

                # Ensure date columns are datetime
                staff_df["Targeted Due Date"] = pd.to_datetime(staff_df["Targeted Due Date"], errors="coerce")
                staff_df["Assigned Date"] = pd.to_datetime(staff_df["Assigned Date"], errors="coerce")

                # --- Top Summary Cards
                col1, col2, col3 = st.columns(3)

                # 1. Total In Progress
                total_in_progress = staff_df.shape[0]

                # 2. Newly Assigned: within last 3 days
                recent_cutoff = datetime.today() - timedelta(days=3)
                newly_assigned = staff_df[staff_df["Assigned Date"] >= recent_cutoff].shape[0]

                # 3. Due within 1 month
                due_soon_cutoff = datetime.today() + timedelta(days=30)
                due_soon = staff_df[staff_df["Targeted Due Date"] <= due_soon_cutoff].shape[0]

                col1.metric("🟡 In Progress", total_in_progress)
                col2.metric("🆕 Newly Assigned (Last 3 days)", newly_assigned)
                col3.metric("📅 Due Within 1 Month", due_soon)
                style_metric_cards(border_left_color="#DBF227")

                # --- Section 1: Mark as Completed
                st.subheader("✅ Mark Requests as Completed")
                # Format dates
                staff_df["Assigned Date"] = staff_df["Assigned Date"].dt.strftime("%Y-%m-%d")
                staff_df["Targeted Due Date"] = staff_df["Targeted Due Date"].dt.strftime("%Y-%m-%d")

                # Display clean table (exclude PriorityOrder column)
                st.dataframe(staff_df[[
                    "Jurisdiction", "Organization", "Name", "Title/Position", "Email Address", "Phone Number",
                    "Focus Area", "TA Type", "Assigned Date", "Targeted Due Date", "Priority", "TA Description"
                ]].reset_index(drop=True))

                # Select request by index (row number in submitted_requests)
                request_indices = staff_df.index.tolist()
                selected_request_index = st.selectbox(
                    "Select a request to marked as completed",
                    options=request_indices,
                    format_func=lambda idx: f"{staff_df.at[idx, 'Name']} | {staff_df.at[idx, 'Jurisdiction']}",
                )
          

                # Submit completion
                if st.button("✅ Mark as Completed"):
                    try:
                        # Map back to original df index
                        global_index = staff_df.loc[selected_request_index].name

                        # Copy + update
                        updated_df = df.copy()
                        updated_df.loc[global_index, "Status"] = "Completed"
                        updated_df.loc[global_index, "Close Date"] = datetime.today().strftime("%Y-%m-%d")

                        # Push to Google Sheet
                        worksheet1.update([updated_df.columns.values.tolist()] + updated_df.values.tolist())

                        st.success("Request marked as completed and synced to Google Sheet.")
                        st.rerun()

                    except Exception as e:
                        st.error(f"Error updating Google Sheets: {str(e)}")

                # --- Submit button styling (CSS injection)
                st.markdown("""
                    <style>
                    .stButton > button {
                        width: 100%;
                        background-color: #cdb4db;
                        color: black;
                        font-weight: 600;
                        border-radius: 8px;
                        padding: 0.6em;
                        margin-top: 1em;
                    }
                    </style>
                """, unsafe_allow_html=True)

                # --- Section 2: Filter, Sort, Comment
                st.subheader("💬 Leave Comments & Track Requests")

                # Filter "In Progress" requests
                if staff_df.empty:
                    st.info("No requests currently in progress.")
                else:
                    # Convert date columns
                    staff_df["Assigned Date"] = pd.to_datetime(staff_df["Assigned Date"], errors="coerce")
                    staff_df["Targeted Due Date"] = pd.to_datetime(staff_df["Targeted Due Date"], errors="coerce")
                    staff_df['Expected Duration (Days)'] = (staff_df["Targeted Due Date"]-staff_df["Assigned Date"]).dt.days

                    # Format dates
                    staff_df["Assigned Date"] = staff_df["Assigned Date"].dt.strftime("%Y-%m-%d")
                    staff_df["Targeted Due Date"] = staff_df["Targeted Due Date"].dt.strftime("%Y-%m-%d")
                    

                    # --- Filters
                    st.markdown("#### 🔍 Filter Options")

                    col1, col2, col3 = st.columns(3)
                    with col1:
                        priority_filter = st.multiselect(
                            "Filter by Priority",
                            options=staff_df["Priority"].unique(),
                            default=staff_df["Priority"].unique(), key='sta1'
                        )

                    with col2:
                        ta_type_filter = st.multiselect(
                            "Filter by TA Type",
                            options=staff_df["TA Type"].unique(),
                            default=staff_df["TA Type"].unique(), key='sta2'
                        )

                    with col3:
                        focus_area_filter = st.multiselect(
                            "Filter by Focus Area",
                            options=staff_df["Focus Area"].unique(),
                            default=staff_df["Focus Area"].unique(), key='sta3'
                        )

                    # Apply filters
                    filtered_df2 = staff_df[
                        (staff_df["Priority"].isin(priority_filter)) &
                        (staff_df["TA Type"].isin(ta_type_filter)) &
                        (staff_df["Focus Area"].isin(focus_area_filter))
                    ]



                    # Display filtered table
                    st.dataframe(filtered_df2[[
                        "Jurisdiction", "Organization", "Name", "Title/Position", "Email Address", "Phone Number",
                        "Focus Area", "TA Type", "Assigned Date", "Targeted Due Date","Expected Duration (Days)","Priority", "Assigned Coach", "TA Description"
                    ]].sort_values(by="Expected Duration (Days)").reset_index(drop=True))

                    # Select request by index (row number in submitted_requests)
                    request_indices2 = filtered_df2.index.tolist()
                    selected_request_index1 = st.selectbox(
                        "Select a request to comment",
                        options=request_indices2,
                        format_func=lambda idx: f"{filtered_df2.at[idx, 'Name']} | {filtered_df2.at[idx, 'Jurisdiction']}",
                    )

                    # Input comment
                    comment_text = st.text_area("Staff Comment", placeholder="Enter comments", height=150, key='commm')

                    # Submit
                    if st.button("✅ Submit Comments"):
                        try:
                            # Get the index of the selected row in the full df
                            global_index = filtered_df2.loc[selected_request_index1].name

                            # Copy df and update
                            updated_df = df.copy()
                            updated_df.loc[global_index, "Staff Comment"] = comment_text
                            updated_df.loc[global_index, "Last Updated"] = datetime.today().strftime("%Y-%m-%d")

                            # Push to Google Sheets
                            worksheet1.update([updated_df.columns.values.tolist()] + updated_df.values.tolist())

                            st.success("💬 Comment saved successfully to 'Staff Comment'.")
                            st.rerun()

                        except Exception as e:
                            st.error(f"Error saving comment: {str(e)}")


