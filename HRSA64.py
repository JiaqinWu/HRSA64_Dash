import streamlit as st
import re
import pandas as pd
from millify import millify # shortens values (10_000 ---> 10k)
from streamlit_extras.metric_cards import style_metric_cards # beautify metric card with css
from datetime import datetime, timedelta
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import altair as alt
import json
import time
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseUpload
import io
from mailjet_rest import Client

st.set_page_config(
        page_title="GU TAP System",
        page_icon="https://raw.githubusercontent.com/JiaqinWu/HRSA64_Dash/main/Georgetown_logo_blueRGB.png", 
        layout="centered"
    ) 

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


def send_email_mailjet(to_email, subject, body):
    api_key = st.secrets["mailjet"]["api_key"]
    api_secret = st.secrets["mailjet"]["api_secret"]
    sender = st.secrets["mailjet"]["sender"]

    mailjet = Client(auth=(api_key, api_secret), version='v3.1')

    data = {
        'Messages': [
            {
                "From": {
                    "Email": sender,
                    "Name": "GU-TAP System"
                },
                "To": [
                    {
                        "Email": to_email,
                        "Name": to_email.split("@")[0]
                    }
                ],
                "Subject": subject,
                "TextPart": body
            }
        ]
    }

    try:
        result = mailjet.send.create(data=data)
        #if result.status_code == 200:
            #st.success(f"📤 Email sent to {to_email}")
        #else:
            #st.warning(f"❌ Failed to email {to_email}: {result.status_code} - {result.json()}")
    except Exception as e:
        st.error(f"❗ Mailjet error: {e}")


def upload_file_to_drive(file, filename, folder_id, creds_dict):
    # Convert Streamlit secret dict into Google Credentials object
    drive_creds = Credentials.from_service_account_info(creds_dict, scopes=[
        "https://www.googleapis.com/auth/drive"
    ])

    # Build the Drive API service
    drive_service = build('drive', 'v3', credentials=drive_creds)

    # Prepare file metadata
    file_metadata = {
        'name': filename,
        'parents': [folder_id]
    }

    # Read uploaded file and prepare media
    media = MediaIoBaseUpload(io.BytesIO(file.read()), mimetype=file.type)

    # Upload file
    uploaded = drive_service.files().create(
        body=file_metadata,
        media_body=media,
        fields='id'
    ).execute()

    # Make the file public
    drive_service.permissions().create(
        fileId=uploaded['id'],
        body={'type': 'anyone', 'role': 'reader'}
    ).execute()

    # Return the sharable link
    return f"https://drive.google.com/file/d/{uploaded['id']}/view"


def _get_records_with_retry(spreadsheet_name, worksheet_name, retries=3, base_delay=0.5):
    """Fetch worksheet records with simple exponential backoff to mitigate 429s."""
    attempt = 0
    last_exc = None
    while attempt < retries:
        try:
            spreadsheet = client.open(spreadsheet_name)
            worksheet = spreadsheet.worksheet(worksheet_name)
            return worksheet.get_all_records()
        except Exception as exc:
            last_exc = exc
            delay = base_delay * (2 ** attempt)
            time.sleep(delay)
            attempt += 1
    # If all retries failed, re-raise last exception
    raise last_exc

@st.cache_data(ttl=600)
def load_main_sheet():
    df = pd.DataFrame(_get_records_with_retry('Example_TA_Request', 'Main'))
    df['Submit Date'] = pd.to_datetime(df['Submit Date'], errors='coerce')
    df["Phone Number"] = df["Phone Number"].astype(str)
    return df

df = load_main_sheet()

# Ensure transfer-related columns exist
transfer_columns = [
    "Last Transfer From",
    "Last Transfer To",
    "Last Transfer Date",
    "Last Transfer By",
    "Transfer History",
]
for _col in transfer_columns:
    if _col not in df.columns:
        df[_col] = ""

# Ensure comment history columns exist
comment_history_columns = [
    "Coordinator Comment History",
    "Staff Comment History",
]
for _col in comment_history_columns:
    if _col not in df.columns:
        df[_col] = ""

@st.cache_data(ttl=600)
def load_interaction_sheet():
    return pd.DataFrame(_get_records_with_retry('Example_TA_Request', 'Interaction'))

df_int = load_interaction_sheet()

# Ensure Interaction sheet has Jurisdiction column for no-ticket logs
if "Jurisdiction" not in df_int.columns:
    df_int["Jurisdiction"] = ""

@st.cache_data(ttl=600)
def load_delivery_sheet():
    return pd.DataFrame(_get_records_with_retry('Example_TA_Request', 'Delivery'))

df_del = load_delivery_sheet()

# Extract last Ticket ID from the existing sheet
last_ticket = df["Ticket ID"].dropna().astype(str).str.extract(r"GU(\d+)", expand=False).astype(int).max()
next_ticket_number = 1 if pd.isna(last_ticket) else last_ticket + 1
new_ticket_id = f"GU{next_ticket_number:04d}"


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
    "jw2104@georgetown.edu": {
        "Coordinator": {"password": "Qin88251216", "name": "Jiaqin Wu"},
        "Assignee/Staff": {"password": "Qin88251216", "name": "Jiaqin Wu"}
    },
    "Jenevieve.Opoku@georgetown.edu": {
        "Coordinator": {"password": "Tootles82!", "name": "Jenevieve Opoku"},
        "Assignee/Staff": {"password": "Tootles82!", "name": "Jenevieve Opoku"}
    },
    "me735@georgetown.edu": {
        "Coordinator": {"password": "me735hrsa64", "name": "Martine Etienne-Mesubi"},
        "Assignee/Staff": {"password": "me735hrsa64", "name": "Martine Etienne-Mesubi"}
    },
    "kd802@georgetown.edu": {
        "Coordinator": {"password": "kd802hrsa!!", "name": "Kemisha Denny"},
        "Assignee/Staff": {"password": "kd802hrsa!!", "name": "Kemisha Denny"}
    },
    "lm1353@georgetown.edu": {
        "Coordinator": {"password": "LM1353hrsa64?", "name": "Lauren Mathae"}
    },
    "katherine.robsky@georgetown.edu": {
        "Coordinator": {"password": "Georgetown1", "name": "Katherine Robsky"},
        "Assignee/Staff": {"password": "Georgetown1", "name": "Katherine Robsky"}
    },
    "db1432@georgetown.edu": {
        "Assignee/Staff": {"password": "Deus123!", "name": "Deus Bazira"}
    },
    "sk2046@georgetown.edu": {
        "Assignee/Staff": {"password": "Sharon123!", "name": "Sharon Kibwana"}
    },
    "sgk23@georgetown.edu": {
        "Assignee/Staff": {"password": "Seble123!", "name": "Seble Kassaye"}
    },
    "weijun.yu@georgetown.edu": {
        "Assignee/Staff": {"password": "Weijun123!", "name": "Weijun Yu"}
    },
    "temesgen.zelalem@mayo.edu": {
        "Assignee/Staff": {"password": "Zelalem123!", "name": "Zelalem Temesgen"}
    },
    "carod@bu.edu": {
        "Assignee/Staff": {"password": "Carlos123!", "name": "Carlos Rodriguez-Diaz"}
    },
    "km2079@georgetown.edu": {
        "Assignee/Staff": {"password": "Kiah123!", "name": "Kiah Moorehead"}
    },
    "vd294@georgetown.edu": {
        "Assignee/Staff": {"password": "Vanessa123!", "name": "Vanessa Da Costa"}
    },
    "tm1649@georgetown.edu": {
        "Assignee/Staff": {"password": "Trena123!", "name": "Trena Mukherjee"}
    },
    "aj1202@georgetown.edu": {
        "Assignee/Staff": {"password": "Abby123!", "name": "Abby Jordan"}
    },
    "mh2504@georgetown.edu": {
        "Assignee/Staff": {"password": "Megan123!", "name": "Megan Highland"}
    },
    "jh2861@georgetown.edu": {
        "Assignee/Staff": {"password": "Jesus123!", "name": "Jesus Hernandez Burgos"}
    },
    "sc2710@georgetown.edu": {
        "Assignee/Staff": {"password": "Samantha123!", "name": "Samantha Cinnick"}
    },
    "bryan.shaw@georgetown.edu": {
        "Assignee/Staff": {"password": "Bryan123!", "name": "Bryan Shaw"}
    },
    "th1089@georgetown.edu": {
        "Assignee/Staff": {"password": "Tara123!", "name": "Tara Hixson"}
    },
    "da988@georgetown.edu":{
        "Assignee/Staff": {"password": "Dzifa123!", "name": "Dzifa Awunyo-Akaba"}
    },
    "mm5674@georgetown.edu":{
        "Assignee/Staff": {"password": "Masill123!", "name": "Masill Miranda"}
    },
    'jb3512@georgetown.edu':{
        "Assignee/Staff": {"password": "Joy123!", "name": "Joy Berry"}
    },
    'ac2992@georgetown.edu':{
        "Assignee/Staff": {"password": "Ashley123!", "name": "Ashley Clonchmore"}
    },
    'gh674@georgetown.edu':{
        "Assignee/Staff": {"password": "Grace123!", "name": "Grace Hazlett"}
    }
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
    #st.image("Georgetown_logo_blueRGB.png",width=200)
    #st.title("Welcome to the GU Technical Assistance Provider System")
    st.markdown(
        """
        <div style='
            display: flex;
            flex-direction: column;
            align-items: center;
            justify-content: center;
            background: #f8f9fa;
            padding: 2em 0 1em 0;
            border-radius: 18px;
            box-shadow: 0 4px 24px rgba(0,0,0,0.07);
            margin-bottom: 2em;
        '>
            <img src='https://raw.githubusercontent.com/JiaqinWu/HRSA64_Dash/main/Georgetown_logo_blueRGB.png' width='200' style='margin-bottom: 1em;'/>
            <h1 style='
                color: #1a237e;
                font-family: "Segoe UI", "Arial", sans-serif;
                font-weight: 700;
                margin: 0;
                font-size: 2.2em;
                text-align: center;
            '>Welcome to the GU Technical Assistance Provider System</h1>
        </div>
        """,
        unsafe_allow_html=True
    )

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
    st.sidebar.markdown(
        f"""
        <div style='
            background: #f8f9fa;
            border-radius: 12px;
            box-shadow: 0 2px 8px rgba(0,0,0,0.04);
            padding: 1.2em 1em 1em 1em;
            margin-bottom: 1.5em;
            text-align: center;
            font-family: Arial, "Segoe UI", sans-serif;
        '>
            <span style='
                font-size: 1.15em;
                font-weight: 700;
                color: #1a237e;
                letter-spacing: 0.5px;
            '>
                Role: {st.session_state.role}
            </span>
        </div>
        """,
        unsafe_allow_html=True
    )

    st.sidebar.button("🔄 Switch Role", on_click=lambda: st.session_state.update({
        "authenticated": False,
        "role": None,
        "user_email": ""
    }))

    # Sidebar: refresh cached datasets
    if st.sidebar.button("🔁 Refresh Data"):
        st.cache_data.clear()
        st.rerun()

    st.markdown("""
        <style>
        .stButton > button {
            width: 100%;
            background-color: #cdb4db;
            color: black;
            font-family: Arial, "Segoe UI", sans-serif;
            font-weight: 600;
            border-radius: 8px;
            padding: 0.6em;
            margin-top: 1em;
            transition: background 0.2s;
        }
        .stButton > button:hover {
            background-color: #b197fc;
            color: #222;
        }
        </style>
    """, unsafe_allow_html=True)

    # Requester: No login needed
    if st.session_state.role == "Requester":
        st.markdown(
            """
            <div style='
                display: flex;
                flex-direction: column;
                align-items: center;
                justify-content: center;
                background: #f8f9fa;
                padding: 2em 0 1em 0;
                border-radius: 18px;
                box-shadow: 0 4px 24px rgba(0,0,0,0.07);
                margin-bottom: 2em;
            '>
                <img src='https://raw.githubusercontent.com/JiaqinWu/HRSA64_Dash/main/Georgetown_logo_blueRGB.png' width='200' style='margin-bottom: 1em;'/>
                <h1 style='
                    color: #1a237e;
                    font-family: "Segoe UI", "Arial", sans-serif;
                    font-weight: 700;
                    margin: 0;
                    font-size: 2.2em;
                    text-align: center;
                '>📥 Georgetown University Technical Assistance Form</h1>
            </div>
            """,
            unsafe_allow_html=True
        )
        #st.header("📥 Georgetown University Technical Assistance Form ")
        st.write("Please complete this form to request Technical Assistance from Georgetown University's Technical Assistance Provider (GU-TAP) team. We will review your request and will be in touch within 1-2 business days. You will receive an email from a TA Coordinator to schedule a time to gather more details about your needs. Once we have this information, we will assign a TA Lead to support you.")
        # Requester form
        col1, col2 = st.columns(2)
        with col1:
            name = st.text_input("Name *",placeholder="Enter text")
        with col2:
            title = st.text_input("Title/Position *",placeholder='Enter text')
        col3, col4 = st.columns(2)
        with col3:
            organization = st.selectbox(
                "Organization *",
                ["GU", "HRSA", "NASTAD"],
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
            focus_area_options = [
                "Housing", "Prevention", "Substance Abuse", "Rapid Start",
                "Telehealth/Telemedicine", "Data Sharing", "Other"
            ]

            focus_area = st.selectbox(
                "TA Focus Area *",
                focus_area_options,
                index=None,
                placeholder="Select option..."
            )

            # If "Other" is selected, show a text input for custom value
            if focus_area == "Other":
                focus_area_other = st.text_input("Please specify the TA Focus Area *")
                if focus_area_other:
                    focus_area = focus_area_other 
        with col8:
            type_TA = st.selectbox(
                "What Style of TA is needed *",
                ["In-Person","Virtual","Hybrid (Combination of in-person and virtual)","Unsure"],
                index=None,
                placeholder="Select option..."
            )
        col9, col10 = st.columns(2)
        with col9:
            due_date = st.date_input(
                "Target Due Date *",
                value=None
            )
            #if not due_date: 
                #st.error("Target Due Date is required.")

            # Add required check: due_date must be after today
            #if due_date and due_date <= datetime.today().date():
                #st.error("Target Due Date must be after today.")

        ta_description = st.text_area("TA Description *", placeholder='Enter text', height=150) 
        document = st.file_uploader(
            "Upload any files or attachments that are relevant to this request.",accept_multiple_files=True
        )
        priority_status = st.selectbox(
                "Priority Status *",
                ["Critical","High","Normal","Low"],
                index=None,
                placeholder="Select option..."
            )

        # Submit button
        st.markdown("""
            <style>
            .stButton > button {
                width: 100%;
                background-color: #cdb4db;
                color: black;
                font-family: Arial, "Segoe UI", sans-serif;
                font-weight: 600;
                border-radius: 8px;
                padding: 0.6em;
                margin-top: 1em;
            }
            </style>
        """, unsafe_allow_html=True)

        # Submit logic
        if st.button("Submit"):
            email_pattern = r'^[\w\.-]+@[\w\.-]+\.\w+$'
            def clean_and_format_us_phone(phone_input):
                digits = re.sub(r'\D', '', phone_input)
                if len(digits) == 10:
                    return f"({digits[:3]}) {digits[3:6]}-{digits[6:]}"
                else:
                    return None

            errors = []

            formatted_phone = clean_and_format_us_phone(phone)


            drive_links = ""
            # Required field checks
            if not name: errors.append("Name is required.")
            if not title: errors.append("Title/Position is required.")
            if not organization: errors.append("Organization must be selected.")
            if not location: errors.append("Location must be selected.")
            if not email or not re.match(email_pattern, email):
                errors.append("Please enter a valid email address.")
            if not phone or not formatted_phone:
                errors.append("Please enter a valid U.S. phone number (10 digits).")
            if not focus_area: errors.append("TA Focus Area must be selected.")
            if not type_TA: errors.append("TA Style must be selected.")
            if not due_date: errors.append("Target Due Date is required.")
            elif due_date <= datetime.today().date():
                errors.append("Target Due Date must be after today.")
            if not ta_description: errors.append("TA Description is required.")
            if not priority_status: errors.append("Priority Status must be selected.")

            # Show warnings or success
            if errors:
                for error in errors:
                    st.warning(error)
            else:
                # Only upload files if all validation passes
                if document:
                    try:
                        folder_id = "1fy1CZSs_t6E6IF68rxblY6YeBihrSPNT" 
                        links = []
                        for file in document:
                            # Rename file as: GU0001_filename.pdf
                            renamed_filename = f"{new_ticket_id}_{file.name}"
                            link = upload_file_to_drive(
                                file=file,
                                filename=renamed_filename,
                                folder_id=folder_id,
                                creds_dict=st.secrets["gcp_service_account"]
                            )
                            links.append(link)
                        drive_links = ", ".join(links)
                        st.success("File(s) uploaded to Google Drive.")    
                    except Exception as e:
                        st.error(f"Error uploading file(s) to Google Drive: {str(e)}")

                new_row = {
                    'Ticket ID': new_ticket_id,
                    'Jurisdiction': location,
                    'Organization': organization,
                    'Name': name,
                    'Title/Position': title,
                    'Email Address': email,
                    "Phone Number": formatted_phone,
                    "Focus Area": focus_area,
                    "TA Type": type_TA,
                    "Targeted Due Date": due_date.strftime("%Y-%m-%d"),
                    "TA Description":ta_description,
                    "Priority": priority_status,
                    "Submit Date": datetime.today().strftime("%Y-%m-%d"),
                    "Status": "Submitted",
                    "Assigned Date": pd.NA,
                    "Close Date": pd.NA,
                    "Assigned Coach": pd.NA,
                    "Coordinator Comment": pd.NA,
                    "Staff Comment": pd.NA,
                    "Document": drive_links
                }
                new_data = pd.DataFrame([new_row])

                try:
                    # Append new data to Google Sheet
                    updated_sheet = pd.concat([df, new_data], ignore_index=True)
                    updated_sheet = updated_sheet.applymap(
                        lambda x: x.strftime("%Y-%m-%d") if isinstance(x, (datetime, pd.Timestamp)) else x
                    )
                    # Replace NaN with empty strings to ensure JSON compatibility
                    updated_sheet = updated_sheet.fillna("")
                    spreadsheet1 = client.open('Example_TA_Request')
                    worksheet1 = spreadsheet1.worksheet('Main')
                    worksheet1.update([updated_sheet.columns.values.tolist()] + updated_sheet.values.tolist())
                    
                    # Clear cache to refresh data
                    st.cache_data.clear()
                    
                    # Send email notifications to all coordinators
                    coordinator_emails = [email for email, user in USERS.items() if "Coordinator" in user]
                    #coordinator_emails = ["jw2104@georgetown.edu"]

                    subject = f"New TA Request Submitted: {new_ticket_id}"
                    for email in coordinator_emails:
                        coordinator_name = USERS[email]["Coordinator"]["name"]
                        personalized_body = f"""
                        Hi {coordinator_name},

                        A new Technical Assistance request has been submitted:

                        Ticket ID: {new_ticket_id}
                        Jurisdiction: {location}
                        Organization: {organization}
                        Name: {name}
                        Description: {ta_description}
                        Priority: {priority_status}
                        Attachments: {drive_links or 'None'}

                        Please review and assign this request via the GU-TAP System: https://hrsagutap.streamlit.app/.
                        Please contact gutap@georgetown.edu for any questions or concerns.

                        Best,
                        GU-TAP System
                        """
                        try:
                            send_email_mailjet(
                                to_email=email,
                                subject=subject,
                                body=personalized_body,
                            )
                        except Exception as e:
                            st.warning(f"⚠️ Failed to send email to coordinator {email}: {e}")


                    # Send confirmation email to requester
                    confirmation_subject = f"Your TA Request ({new_ticket_id}) has been received"
                    confirmation_body = f"""
                    Hi {name},

                    Thank you for submitting your Technical Assistance request to Georgetown University's Technical Assistance Provider (GU-TAP) team.

                    Here is a summary of your submission:
                    - Ticket ID: {new_ticket_id}
                    - Jurisdiction: {location}
                    - Organization: {organization}
                    - Name: {name}
                    - Title/Position: {title}
                    - Email Address: {email}
                    - Phone Number: {formatted_phone}
                    - Focus Area: {focus_area}
                    - TA Type: {type_TA}
                    - Targeted Due Date: {due_date.strftime("%Y-%m-%d")}
                    - Priority: {priority_status}
                    - Description: {ta_description}

                    A TA Coordinator will review your request and assign a coach to your request within 2-3 business days.
                    Please contact gutap@georgetown.edu for any questions or concerns.

                    Best,
                    GU-TAP System
                    """

                    try:
                        send_email_mailjet(
                            to_email=email,
                            subject=confirmation_subject,
                            body=confirmation_body,
                        )
                    except Exception as e:
                        st.warning(f"⚠️ Failed to send confirmation email to requester: {e}")


                    st.success("✅ Submission successful!")
                    
                    # Clear cache to refresh data
                    st.cache_data.clear()
                    
                    # Wait a moment then redirect to main page
                    time.sleep(3)
                    st.rerun()

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
                user_roles = USERS.get(email)
                if user_roles and st.session_state.role in user_roles:
                    user = user_roles[st.session_state.role]
                    if user["password"] == password:
                        st.session_state.authenticated = True
                        st.session_state.user_email = email
                        st.success("Login successful!")
                        st.rerun()
                    else:
                        st.error("Invalid credentials or role mismatch.")
                else:
                    st.error("Invalid credentials or role mismatch.")

        else:
            if st.session_state.role == "Coordinator":
                user_info = USERS.get(st.session_state.user_email)
                coordinator_name = user_info["Coordinator"]["name"]
                st.markdown(
                    """
                    <div style='
                        display: flex;
                        flex-direction: column;
                        align-items: center;
                        justify-content: center;
                        background: #f8f9fa;
                        padding: 2em 0 1em 0;
                        border-radius: 18px;
                        box-shadow: 0 4px 24px rgba(0,0,0,0.07);
                        margin-bottom: 2em;
                    '>
                        <img src='https://raw.githubusercontent.com/JiaqinWu/HRSA64_Dash/main/Georgetown_logo_blueRGB.png' width='200' style='margin-bottom: 1em;'/>
                        <h1 style='
                            color: #1a237e;
                            font-family: "Segoe UI", "Arial", sans-serif;
                            font-weight: 700;
                            margin: 0;
                            font-size: 2.2em;
                            text-align: center;
                        '>📬 Coordinator Dashboard</h1>
                    </div>
                    """,
                    unsafe_allow_html=True
                )
                #st.header("📬 Coordinator Dashboard")
                # Personalized greeting
                if user_info and "Coordinator" in user_info:
                    st.markdown(f"""
                    <div style='                      
                    background: #f8f9fa;                        
                    border-radius: 12px;                        
                    box-shadow: 0 2px 8px rgba(0,0,0,0.04);                        
                    padding: 1.2em 1em 1em 1em;                        
                    margin-bottom: 1.5em;                        
                    text-align: center;                        
                    font-family: Arial, "Segoe UI", sans-serif;                    
                    '>
                        <span style='                           
                        font-size: 1.15em;
                        font-weight: 700;
                        color: #1a237e;
                        letter-spacing: 0.5px;'>
                            👋 Welcome, {coordinator_name}!
                        </span>
                    </div>
                    """, unsafe_allow_html=True)
                col1, col2, col3 = st.columns(3)
                total_request = df['Ticket ID'].nunique()
                inprogress_request = df[df['Status'] == 'In Progress']['Ticket ID'].nunique()
                completed_request = df[df['Status'] == 'Completed']['Ticket ID'].nunique()

                col1.metric(label="# of Total Requests", value= millify(total_request, precision=2))
                col2.metric(label="# of In-Progress Requests", value= millify(inprogress_request, precision=2))
                col3.metric(label="# of Completed Requests", value= millify(completed_request, precision=2))
                style_metric_cards(border_left_color="#DBF227")

                col1, col2, col3 = st.columns(3)
                # create column span
                today = datetime.today()
                last_week = today - timedelta(days=7)
                last_month = today - timedelta(days=30)
                undone_request = df[df['Status'] == 'Submitted']['Ticket ID'].nunique()
                pastweek_request = df[df['Submit Date'] >= last_week]['Ticket ID'].nunique()
                pastmonth_request = df[df['Submit Date'] >= last_month]['Ticket ID'].nunique()
                col1.metric(label="# of Unassigned Requests", value= millify(undone_request, precision=2))
                col2.metric(label="# of Requests from past week", value= millify(pastweek_request, precision=2))
                col3.metric(label="# of Requests from past month", value= millify(pastmonth_request, precision=2))
                style_metric_cards(border_left_color="#DBF227")
                with st.expander("🔎 **MONITOR IN-PROGRESS REQUESTS**"):
                    st.markdown("""
                        <div style='
                            background: #f0f4ff;
                            border-radius: 16px;
                            box-shadow: 0 2px 8px rgba(26,35,126,0.08);
                            padding: 1.5em 1em 1em 1em;
                            margin-bottom: 2em;
                            margin-top: 1em;
                        '>
                            <div style='
                                color: #1a237e;
                                font-family: "Segoe UI", "Arial", sans-serif;
                                font-weight: 700;
                                font-size: 1.4em;
                                margin-bottom: 0.3em;
                            '>📊 Monitor In-Progress TA Requests</div>
                            <div style='
                                color: #333;
                                font-size: 1.08em;
                                margin-bottom: 0.8em;
                                Track all active Technical Assistance requests, view staff assignments, and monitor upcoming deadlines. Use the interactive charts and filters below to stay on top of your team's workload.
                            </div>
                        </div>
                    """, unsafe_allow_html=True)
                    # Convert date columns
                    df["Assigned Date"] = pd.to_datetime(df["Assigned Date"], errors="coerce")
                    df["Targeted Due Date"] = pd.to_datetime(df["Targeted Due Date"], errors="coerce")
                    df["Submit Date"] = pd.to_datetime(df["Submit Date"], errors="coerce")

                    inprogress = df[df['Status'] == 'In Progress']
                    next_month = today + timedelta(days=30)
                    request_pastmonth = df[(df["Targeted Due Date"] <= next_month)&(df['Status'] == 'In Progress')]

                    col4, col5 = st.columns(2)
                    # --- Pie 1: In Progress
                    with col4:
                        st.markdown("##### 🟡 In Progress Requests by Coach")
                        if not inprogress.empty:
                            chart_data = inprogress['Assigned Coach'].value_counts().reset_index()
                            chart_data.columns = ['Assigned Coach', 'Count']
                            pie1 = alt.Chart(chart_data).mark_arc(innerRadius=50).encode(
                                theta=alt.Theta(field="Count", type="quantitative"),
                                color=alt.Color(field="Assigned Coach", type="nominal"),
                                tooltip=["Assigned Coach", "Count"]
                            ).properties(width=250, height=250)
                            st.altair_chart(pie1, use_container_width=True)
                        else:
                            st.info("No in-progress requests to show.")

                    # --- Pie 2: Requests in Past Month
                    with col5:
                        st.markdown("##### 📅 Due in 30 Days by Coach")
                        if not request_pastmonth.empty:
                            chart_data = request_pastmonth['Assigned Coach'].value_counts().reset_index()
                            chart_data.columns = ['Assigned Coach', 'Count']
                            pie2 = alt.Chart(chart_data).mark_arc(innerRadius=50).encode(
                                theta=alt.Theta(field="Count", type="quantitative"),
                                color=alt.Color(field="Assigned Coach", type="nominal"),
                                tooltip=["Assigned Coach", "Count"]
                            ).properties(width=250, height=250)
                            st.altair_chart(pie2, use_container_width=True)
                        else:
                            st.info("No requests with coming dues to show.")


                    staff_list = ["Jenevieve Opoku", "Deus Bazira", "Kemisha Denny", "Katherine Robsky", 
                    "Martine Etienne-Mesubi", "Seble Kassaye", "Weijun Yu", "Jiaqin Wu", "Zelalem Temesgen", "Carlos Rodriguez-Diaz",
                    "Kiah Moorehead","Vanessa Da Costa","Trena Mukherjee","Abby Jordan","Megan Highland","Jesus Hernandez Burgos",
                    "Samantha Cinnick","Bryan Shaw","Tara Hixson","Dzifa Awunyo-Akaba","Masill Miranda","Joy Berry","Ashley Clonchmore",
                    "Grace Hazlett"]

                    staff_list_sorted = sorted(staff_list, key=lambda x: x.split()[0])

                    selected_staff = st.selectbox("Select a staff to view their requests", staff_list_sorted, index=None,
                            placeholder="Select option...")

                    today = datetime.today()
                    last_month = today - timedelta(days=30)
                    staff_dff = df[df["Assigned Coach"] == selected_staff].copy()
                    in_progress_count = staff_dff[staff_dff["Status"] == "In Progress"]['Ticket ID'].nunique()
                    due_soon_count = staff_dff[
                        (staff_dff["Status"] == "In Progress") & 
                        (staff_dff["Targeted Due Date"] <= next_month)
                    ]['Ticket ID'].nunique()
                    completed_recently = staff_dff[
                        (staff_dff["Status"] == "Completed") & 
                        (staff_dff["Submit Date"] >= last_month)
                    ]['Ticket ID'].nunique()

                    # Metric 
                    col1, col2, col3 = st.columns(3)
                    col1.metric("🟡 In Progress", in_progress_count)
                    col2.metric("📅 Due in 30 Days", due_soon_count)
                    col3.metric("✅ Completed (Last 30 Days)", completed_recently)

                    # Detailed Table
                    st.markdown("##### 📋 Detailed Request List")

                    # Status filter
                    status_options = ["In Progress", "Completed"]
                    selected_status = st.multiselect(
                        "Filter by request status",
                        options=status_options,
                        default=["In Progress"]
                    )

                    # Apply filters: staff and status
                    staff_dfff = staff_dff[staff_dff["Status"].isin(selected_status)].copy()

                    display_cols = [
                        "Ticket ID", "Jurisdiction", "Organization", "Name", "Focus Area", "TA Type",
                        "Targeted Due Date", "Priority", "Status", "TA Description","Document"
                    ]

                    # Sort by due date
                    staff_dfff = staff_dfff.sort_values(by="Targeted Due Date")

                    # Format dates for display
                    staff_dfff["Targeted Due Date"] = staff_dfff["Targeted Due Date"].dt.strftime("%Y-%m-%d")

                    st.dataframe(staff_dfff[display_cols].reset_index(drop=True))

                    # Filter submitted requests
                    submitted_requests = df[df["Status"] == "Submitted"].copy()

                st.markdown("<hr style='margin:2em 0; border:1px solid #dee2e6;'>", unsafe_allow_html=True)
                with st.expander("📝 **ASSIGN TA REQUESTS**"):
                    st.markdown("""
                        <div style='background: #f0f4ff; border-radius: 16px; box-shadow: 0 2px 8px rgba(26,35,126,0.08); padding: 1.5em 1em 1em 1em; margin-bottom: 2em; margin-top: 1em;'>
                            <div style='color: #1a237e; font-family: "Segoe UI", "Arial", sans-serif; font-weight: 700; font-size: 1.4em; margin-bottom: 0.3em;'>📋 Assign TA Requests</div>
                            <div style='color: #333; font-size: 1.08em; margin-bottom: 0.8em;'>
                                Review all unassigned Technical Assistance requests and assign them to the appropriate staff member. Use the table and filters below to prioritize and manage new requests efficiently.
                            </div>
                        </div>
                    """, unsafe_allow_html=True)
                    st.markdown("#### 📋 Unassigned Requests")

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
                            "Ticket ID","Jurisdiction", "Organization", "Name", "Title/Position", "Email Address", "Phone Number",
                            "Focus Area", "TA Type", "Submit Date", "Targeted Due Date", "Priority", "TA Description","Document"
                        ]].reset_index(drop=True))

                        # Select request by index 
                        request_indices = submitted_requests_sorted.index.tolist()
                        selected_request_index = st.selectbox(
                            "Select a request to assign",
                            options=request_indices,
                            format_func=lambda idx: f"{submitted_requests_sorted.at[idx, 'Ticket ID']} | {submitted_requests_sorted.at[idx, 'Name']} | {submitted_requests_sorted.at[idx, 'Jurisdiction']}",
                        )

                        # Select coach
                        selected_coach = st.selectbox(
                            "Assign a coach",
                            options=staff_list_sorted,
                            index=None,
                            placeholder="Select option..."
                        )

                        # Assign button
                        if st.button("✅ Assign Coach and Start TA"):
                            try:
                                updated_df = df.copy()
                                # Update the selected row
                                updated_df.loc[selected_request_index, "Assigned Coach"] = selected_coach
                                updated_df.loc[selected_request_index, "Assigned Coordinator"] = coordinator_name
                                updated_df.loc[selected_request_index, "Status"] = "In Progress"
                                updated_df.loc[selected_request_index, "Assigned Date"] = datetime.today().strftime("%Y-%m-%d")

                                updated_df = updated_df.applymap(
                                    lambda x: x.strftime("%Y-%m-%d") if isinstance(x, (pd.Timestamp, datetime)) and not pd.isna(x) else x
                                )
                                updated_df = updated_df.fillna("") 
                                spreadsheet1 = client.open('Example_TA_Request')
                                worksheet1 = spreadsheet1.worksheet('Main')

                                # Push to Google Sheet
                                worksheet1.update([updated_df.columns.values.tolist()] + updated_df.values.tolist())

                                # Clear cache to refresh data
                                st.cache_data.clear()
                                
                                st.success(f"Coach {selected_coach} assigned! Status updated to 'In Progress'.")

                                # Send email to staff   
                                # Find staff email by name
                                staff_email = None
                                for email, roles in USERS.items():
                                    if "Assignee/Staff" in roles and roles["Assignee/Staff"]["name"] == selected_coach:
                                        staff_email = email
                                        break

                                if staff_email:
                                    staff_subject = f"You have been assigned a new TA request: {updated_df.loc[selected_request_index, 'Ticket ID']}"
                                    staff_body = f"""
                                    Hi {selected_coach},

                                    You have been assigned as the coach for the following Technical Assistance request:

                                    Ticket ID: {updated_df.loc[selected_request_index, 'Ticket ID']}
                                    Jurisdiction: {updated_df.loc[selected_request_index, 'Jurisdiction']}
                                    Organization: {updated_df.loc[selected_request_index, 'Organization']}
                                    Name: {updated_df.loc[selected_request_index, 'Name']}
                                    Description: {updated_df.loc[selected_request_index, 'TA Description']}
                                    Priority: {updated_df.loc[selected_request_index, 'Priority']}
                                    Targeted Due Date: {updated_df.loc[selected_request_index, 'Targeted Due Date']}
                                    Attachments: {updated_df.loc[selected_request_index, 'Document'] or 'None'}

                                    Please view and manage this request via the GU-TAP System: https://hrsagutap.streamlit.app/.
                                    Please contact gutap@georgetown.edu for any questions or concerns.

                                    Best,
                                    GU-TAP System
                                    """
                                    try:
                                        send_email_mailjet(
                                            to_email=staff_email,
                                            subject=staff_subject,
                                            body=staff_body,
                                        )
                                    except Exception as e:
                                        st.warning(f"⚠️ Failed to send assignment email to staff {selected_coach}: {e}")

                                time.sleep(2)
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

                # --- Transfer TA Requests (Coordinator only)
                st.markdown("<hr style='margin:2em 0; border:1px solid #dee2e6;'>", unsafe_allow_html=True)
                with st.expander("🔄 **TRANSFER TA REQUESTS**"):
                    st.markdown("""
                        <div style='background: #f0f4ff; border-radius: 16px; box-shadow: 0 2px 8px rgba(26,35,126,0.08); padding: 1.5em 1em 1em 1em; margin-bottom: 2em; margin-top: 1em;'>
                            <div style='color: #1a237e; font-family: "Segoe UI", "Arial", sans-serif; font-weight: 700; font-size: 1.4em; margin-bottom: 0.3em;'>🔄 Transfer In-Progress TA Requests</div>
                            <div style='color: #333; font-size: 1.08em; margin-bottom: 0.8em;'>
                                Reassign an in-progress request to another coach if needed.
                            </div>
                        </div>
                    """, unsafe_allow_html=True)

                    inprogress_for_transfer = df[df['Status'] == 'In Progress'].copy()

                    if inprogress_for_transfer.empty:
                        st.info("No in-progress requests available to transfer.")
                    else:
                        # Selection of ticket to transfer
                        transfer_indices = inprogress_for_transfer.index.tolist()
                        selected_transfer_index = st.selectbox(
                            "Select a request to transfer",
                            options=transfer_indices,
                            format_func=lambda idx: f"{df.at[idx, 'Ticket ID']} | {df.at[idx, 'Jurisdiction']} (Current: {df.at[idx, 'Assigned Coach']})",
                        )

                        current_coach = df.at[selected_transfer_index, 'Assigned Coach']
                        # Exclude current coach from options
                        available_new_coaches = [c for c in staff_list_sorted if c != current_coach]
                        new_coach = st.selectbox(
                            "Transfer to coach",
                            options=available_new_coaches,
                            index=None,
                            placeholder="Select option..."
                        )

                        reason_transfer = st.text_area("Reason for transfer (optional)", placeholder="Enter text", height=100)

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

                        if st.button("✅ Confirm Transfer"):
                            if not new_coach:
                                st.warning("Please select a new coach to transfer to.")
                            else:
                                try:
                                    updated_df = df.copy()
                                    old_coach = current_coach

                                    # Update transfer metadata
                                    updated_df.loc[selected_transfer_index, "Last Transfer From"] = old_coach or ""
                                    updated_df.loc[selected_transfer_index, "Last Transfer To"] = new_coach
                                    updated_df.loc[selected_transfer_index, "Last Transfer Date"] = datetime.today().strftime("%Y-%m-%d")
                                    updated_df.loc[selected_transfer_index, "Last Transfer By"] = coordinator_name
                                    updated_df.loc[selected_transfer_index, "Reason for Transfer"] = reason_transfer or ""

                                    # Update active assignment
                                    updated_df.loc[selected_transfer_index, "Assigned Coach"] = new_coach

                                    # Append transfer history
                                    history_entry = f"{datetime.today().strftime('%Y-%m-%d')} | By: {coordinator_name} | From: {old_coach or 'N/A'} -> To: {new_coach}" + (f" | Reason: {reason_transfer.strip()}" if reason_transfer else "")
                                    existing_history = str(updated_df.loc[selected_transfer_index, "Transfer History"]).strip()
                                    if existing_history and existing_history.lower() != "nan":
                                        updated_df.loc[selected_transfer_index, "Transfer History"] = existing_history + "\n" + history_entry
                                    else:
                                        updated_df.loc[selected_transfer_index, "Transfer History"] = history_entry

                                    # Stringify dates for sheet
                                    updated_df = updated_df.applymap(
                                        lambda x: x.strftime("%Y-%m-%d") if isinstance(x, (pd.Timestamp, datetime)) and not pd.isna(x) else x
                                    )
                                    updated_df = updated_df.fillna("")

                                    spreadsheet1 = client.open('Example_TA_Request')
                                    worksheet1 = spreadsheet1.worksheet('Main')
                                    worksheet1.update([updated_df.columns.values.tolist()] + updated_df.values.tolist())

                                    st.cache_data.clear()

                                    st.success(f"Request {updated_df.loc[selected_transfer_index, 'Ticket ID']} transferred from {old_coach or 'N/A'} to {new_coach}.")

                                    # Email notification to new assignee
                                    staff_email = None
                                    for _email, roles in USERS.items():
                                        if "Assignee/Staff" in roles and roles["Assignee/Staff"]["name"] == new_coach:
                                            staff_email = _email
                                            break

                                    if staff_email:
                                        staff_subject = f"You have been transferred a TA request: {updated_df.loc[selected_transfer_index, 'Ticket ID']}"
                                        staff_body = f"""
                                        Hi {new_coach},

                                        You have been assigned (via transfer) as the coach for the following Technical Assistance request:

                                        Ticket ID: {updated_df.loc[selected_transfer_index, 'Ticket ID']}
                                        Jurisdiction: {updated_df.loc[selected_transfer_index, 'Jurisdiction']}
                                        Organization: {updated_df.loc[selected_transfer_index, 'Organization']}
                                        Previous Coach: {old_coach or 'N/A'}
                                        Description: {updated_df.loc[selected_transfer_index, 'TA Description']}
                                        Priority: {updated_df.loc[selected_transfer_index, 'Priority']}
                                        Targeted Due Date: {updated_df.loc[selected_transfer_index, 'Targeted Due Date']}
                                        Attachments: {updated_df.loc[selected_transfer_index, 'Document'] or 'None'}

                                        Decision by: {coordinator_name}
                                        {('Reason: ' + reason_transfer) if reason_transfer else ''}

                                        Please view and manage this request via the GU-TAP System: https://hrsagutap.streamlit.app/.
                                        Please contact gutap@georgetown.edu for any questions or concerns.

                                        Best,
                                        GU-TAP System
                                        """
                                        try:
                                            send_email_mailjet(
                                                to_email=staff_email,
                                                subject=staff_subject,
                                                body=staff_body,
                                            )
                                        except Exception as e:
                                            st.warning(f"⚠️ Failed to send transfer email to staff {new_coach}: {e}")

                                    # Email notification to previous coach (if exists)
                                    if old_coach and old_coach != new_coach:
                                        previous_coach_email = None
                                        for _email, roles in USERS.items():
                                            if "Assignee/Staff" in roles and roles["Assignee/Staff"]["name"] == old_coach:
                                                previous_coach_email = _email
                                                break

                                        if previous_coach_email:
                                            previous_subject = f"TA Request Transferred: {updated_df.loc[selected_transfer_index, 'Ticket ID']}"
                                            previous_body = f"""
                                            Hi {old_coach},

                                            The following Technical Assistance request has been transferred from you to another coach:

                                            Ticket ID: {updated_df.loc[selected_transfer_index, 'Ticket ID']}
                                            Jurisdiction: {updated_df.loc[selected_transfer_index, 'Jurisdiction']}
                                            Organization: {updated_df.loc[selected_transfer_index, 'Organization']}
                                            New Coach: {new_coach}
                                            Description: {updated_df.loc[selected_transfer_index, 'TA Description']}
                                            Priority: {updated_df.loc[selected_transfer_index, 'Priority']}
                                            Targeted Due Date: {updated_df.loc[selected_transfer_index, 'Targeted Due Date']}

                                            Decision by: {coordinator_name}
                                            {('Reason: ' + reason_transfer) if reason_transfer else ''}

                                            You no longer need to work on this request. Please contact gutap@georgetown.edu for any questions or concerns.

                                            Best,
                                            GU-TAP System
                                            """
                                            try:
                                                send_email_mailjet(
                                                    to_email=previous_coach_email,
                                                    subject=previous_subject,
                                                    body=previous_body,
                                                )
                                            except Exception as e:
                                                st.warning(f"⚠️ Failed to send transfer notification to previous coach {old_coach}: {e}")

                                    time.sleep(2)
                                    st.rerun()
                                except Exception as e:
                                    st.error(f"Error updating Google Sheets: {str(e)}")

                st.markdown("<hr style='margin:2em 0; border:1px solid #dee2e6;'>", unsafe_allow_html=True)
                with st.expander("👍 **DETAILS OF IN-PROGRESS & COMPLETED REQUESTS**"):
                    st.markdown("""
                        <div style='
                            background: #e9ecef;
                            border-radius: 14px;
                            box-shadow: 0 2px 8px rgba(0,0,0,0.04);
                            padding: 1.5em 1em 1em 1em;
                            margin-bottom: 2em;
                            margin-top: 1em;
                        '>
                            <h2 style='
                                color: #1a237e;
                                font-family: "Segoe UI", "Arial", sans-serif;
                                font-weight: 700;
                                margin-bottom: 0.2em;
                                font-size: 1.3em;
                            '>📊 TA Request Management: Comments & Completion Review</h2>
                            <p style='
                                color: #333;
                                font-size: 1em;
                                margin-bottom: 0.8em;
                            '>
                                Use this section to leave comments or updates for in-progress TA requests, and to review the status and details of completed requests.
                            </p>
                        </div>
                    """, unsafe_allow_html=True)

                    st.markdown("#### 🚧 In-progress Requests")


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
                        st.markdown("##### 🔍 Filter Options")

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
                            "Ticket ID","Jurisdiction", "Organization", "Name", "Title/Position", "Email Address", "Phone Number",
                            "Focus Area", "TA Type", "Assigned Date", "Targeted Due Date","Expected Duration (Days)","Priority",
                            "Assigned Coach", "TA Description","Document","Coordinator Comment History"
                        ]].sort_values(by="Expected Duration (Days)").reset_index(drop=True))

                        # Select request by index (row number in submitted_requests)
                        request_indices1 = filtered_df.index.tolist()
                        selected_request_index1 = st.selectbox(
                            "Select a request to comment",
                            options=request_indices1,
                            format_func=lambda idx: f"{filtered_df.at[idx, 'Ticket ID']} | {filtered_df.at[idx, 'Name']} | {filtered_df.at[idx, 'Jurisdiction']}",
                        )

                        # Input + submit comment
                        comment_input = st.text_area("Comments", placeholder="Enter comments", height=150, key='comm')

                        if st.button("✅ Submit Comments"):
                            try:
                                # Find the actual row index in the original df (map back using ID or index)
                                selected_row_global_index = filtered_df.loc[selected_request_index1].name

                                # Copy and update df
                                updated_df = df.copy()
                                # Keep latest comment in main field
                                updated_df.loc[selected_row_global_index, "Coordinator Comment"] = comment_input
                                # Append to history with timestamp and author
                                ts = datetime.today().strftime("%Y-%m-%d %H:%M")
                                author = coordinator_name
                                entry = f"{ts} | {author}: {comment_input}" if comment_input else ""
                                if entry:
                                    existing = str(updated_df.loc[selected_row_global_index, "Coordinator Comment History"]).strip()
                                    if existing and existing.lower() != "nan":
                                        updated_df.loc[selected_row_global_index, "Coordinator Comment History"] = existing + "\n" + entry
                                    else:
                                        updated_df.loc[selected_row_global_index, "Coordinator Comment History"] = entry
                                updated_df = updated_df.applymap(
                                    lambda x: x.strftime("%Y-%m-%d") if isinstance(x, (pd.Timestamp, datetime)) and not pd.isna(x) else x
                                )
                                updated_df = updated_df.fillna("") 

                                spreadsheet1 = client.open('Example_TA_Request')
                                worksheet1 = spreadsheet1.worksheet('Main')

                                # Push to Google Sheets
                                spreadsheet1 = client.open('Example_TA_Request')
                                worksheet1 = spreadsheet1.worksheet('Main')
                                worksheet1.update([updated_df.columns.values.tolist()] + updated_df.values.tolist())

                                # Clear cache to refresh data
                                st.cache_data.clear()
                                
                                st.success("💬 Comment saved successfully!.")
                                time.sleep(2)
                                st.rerun()

                            except Exception as e:
                                st.error(f"Error updating Google Sheets: {str(e)}")

                    st.markdown("#### ✅ Completed Requests")


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
                        st.markdown("##### 🔍 Filter Options")

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
                            "Ticket ID","Jurisdiction", "Organization", "Name", "Title/Position", "Email Address", "Phone Number",
                            "Focus Area", "TA Type", "Priority", "Assigned Coach", "TA Description","Document","Assigned Date",
                            "Targeted Due Date", "Close Date", "Expected Duration (Days)",
                            'Actual Duration (Days)', "Coordinator Comment History", "Staff Comment History", "Transfer History"
                        ]].reset_index(drop=True))

                st.markdown("<hr style='margin:2em 0; border:1px solid #dee2e6;'>", unsafe_allow_html=True)
                with st.expander("🗒️ **SUBMIT INTERACTION LOG**"):
                    st.markdown("""
                        <div style='background: #f0f4ff; border-radius: 16px; box-shadow: 0 2px 8px rgba(26,35,126,0.08); padding: 1.5em 1em 1em 1em; margin-bottom: 2em; margin-top: 1em;'>
                            <div style='color: #1a237e; font-family: "Segoe UI", "Arial", sans-serif; font-weight: 700; font-size: 1.4em; margin-bottom: 0.3em;'>🗒️ Submit a New Interaction Log Form</div>
                            <div style='color: #333; font-size: 1.08em; margin-bottom: 0.8em;'>
                                Log a new interaction with a jurisdiction. Fill out the form below to record emails, meetings, or other communications related to a TA request.
                            </div>
                        </div>
                    """, unsafe_allow_html=True)
                    lis_ticket = ["No Ticket ID"] + sorted([tid for tid in df["Ticket ID"].dropna().astype(str).unique().tolist()])

                    # Interaction Log form
                    col1, col2 = st.columns(2)
                    with col1:
                        ticket_id_int = st.selectbox("Ticket ID *", lis_ticket, index=None,
                            placeholder="Select option...", key='interaction')
                    with col2:
                        date_int = st.date_input("Date of Interaction *",value=datetime.today().date())

                     # If No Ticket ID, ask for Jurisdiction
                    jurisdiction_for_no_ticket = None
                    if ticket_id_int == "No Ticket ID":
                        jurisdiction_for_no_ticket = st.selectbox(
                            "Jurisdiction *",
                            lis_location,
                            index=None,
                            placeholder="Select option...",
                            key='juris_interaction'
                        )

                    list_interaction = [
                        "Email", "Phone Call", "In-Person Meeting", "Online Meeting", "Other"
                    ]

                    type_interaction = st.selectbox(
                        "Type of Interaction *",
                        list_interaction,
                        index=None,
                        placeholder="Select option..."
                    )

                    # If "Other" is selected, show a text input for custom value
                    if type_interaction == "Other":
                        type_interaction_other = st.text_input("Please specify the Type of Interaction *")
                        if type_interaction_other:
                            type_interaction = type_interaction_other 
                    interaction_description = st.text_area("Short Summary *", placeholder='Enter text', height=150,key='interaction_description') 

                   
                    document_int = st.file_uploader(
                        "Upload any files or attachments that are relevant to this interaction.",accept_multiple_files=True
                    )

                    # Submit button
                    st.markdown("""
                        <style>
                        .stButton > button {
                            width: 100%;
                            background-color: #cdb4db;
                            color: black;
                            font-family: Arial, "Segoe UI", sans-serif;
                            font-weight: 600;
                            border-radius: 8px;
                            padding: 0.6em;
                            margin-top: 1em;
                        }
                        </style>
                    """, unsafe_allow_html=True)

                    # Submit logic
                    if st.button("Submit",key='interaction_submit'):
                        errors = []
                        drive_links_int = ""  # Initialize here
                        # Required field checks
                        if not ticket_id_int: errors.append("Ticket ID is required.")
                        if ticket_id_int == "No Ticket ID" and not jurisdiction_for_no_ticket:
                            errors.append("Jurisdiction is required when Ticket ID is not provided.")
                        if not date_int: errors.append("Date of interaction is required.")
                        if not type_interaction: errors.append("Type of interaction is required.")
                        if not interaction_description: errors.append("Short summary is required.")

                        # Show warnings or success
                        if errors:
                            for error in errors:
                                st.warning(error)
                        else:
                            # Only upload files if all validation passes
                            if document_int:
                                try:
                                    folder_id_int = "19-Sm8W151tg1zyDN0Nh14DUvOVUieqq7" 
                                    links_int = []
                                    for file in document_int:
                                        # Rename file as: GU0001_filename.pdf
                                        renamed_filename = f"{ticket_id_int}_{file.name}"
                                        link = upload_file_to_drive(
                                            file=file,
                                            filename=renamed_filename,
                                            folder_id=folder_id_int,
                                            creds_dict=st.secrets["gcp_service_account"]
                                        )
                                        links_int.append(link)
                                    drive_links_int = ", ".join(links_int)
                                    st.success("File(s) uploaded to Google Drive.")    
                                except Exception as e:
                                    st.error(f"Error uploading file(s) to Google Drive: {str(e)}")

                            new_row_int = {
                                'Ticket ID': ticket_id_int,
                                "Date of Interaction": date_int.strftime("%Y-%m-%d"),
                                "Type of Interaction": type_interaction,
                                "Short Summary": interaction_description,
                                "Document": drive_links_int,
                                "Jurisdiction": (lambda: (
                                    # If a Ticket ID is provided, extract jurisdiction from Main sheet
                                    str(df.loc[df["Ticket ID"].astype(str) == str(ticket_id_int), "Jurisdiction"].iloc[0])
                                    if ticket_id_int != "No Ticket ID" and not df.loc[df["Ticket ID"].astype(str) == str(ticket_id_int), "Jurisdiction"].empty
                                    else (jurisdiction_for_no_ticket or "")
                                ))(),
                                "Submitted By": coordinator_name,
                                "Submission Date": datetime.today().strftime("%Y-%m-%d %H:%M")
                            }
                            new_data_int = pd.DataFrame([new_row_int])

                            try:
                                # Append new data to Google Sheet
                                updated_sheet1 = pd.concat([df_int, new_data_int], ignore_index=True)
                                updated_sheet1= updated_sheet1.applymap(
                                    lambda x: x.strftime("%Y-%m-%d") if isinstance(x, (datetime, pd.Timestamp)) else x
                                )
                                # Replace NaN with empty strings to ensure JSON compatibility
                                updated_sheet1 = updated_sheet1.fillna("")
                                spreadsheet2 = client.open('Example_TA_Request')
                                worksheet2 = spreadsheet2.worksheet('Interaction')
                                worksheet2.update([updated_sheet1.columns.values.tolist()] + updated_sheet1.values.tolist())

                                # Clear cache to refresh data
                                st.cache_data.clear()
                                
                                st.success("✅ Submission successful!")
                                time.sleep(2)
                                st.rerun()

                            except Exception as e:
                                st.error(f"Error updating Google Sheets: {str(e)}")

                st.markdown("<hr style='margin:2em 0; border:1px solid #dee2e6;'>", unsafe_allow_html=True)

                with st.expander("📦 **SUBMIT DELIVERY FORM**"):
                    st.markdown("""
                        <div style='background: #f0f4ff; border-radius: 16px; box-shadow: 0 2px 8px rgba(26,35,126,0.08); padding: 1.5em 1em 1em 1em; margin-bottom: 2em; margin-top: 1em;'>
                            <div style='color: #1a237e; font-family: "Segoe UI", "Arial", sans-serif; font-weight: 700; font-size: 1.4em; margin-bottom: 0.3em;'>📦 Submit a New Delivery Form</div>
                            <div style='color: #333; font-size: 1.08em; margin-bottom: 0.8em;'>
                                Record a new delivery (e.g., report, dashboard, data) for a TA request. Use the form below to upload files and provide a summary of the delivery.
                            </div>
                        </div>
                    """, unsafe_allow_html=True)
                    lis_ticket1 = df["Ticket ID"].unique().tolist()

                    # Interaction Log form
                    col1, col2 = st.columns(2)
                    with col1:
                        ticket_id_del = st.selectbox("Ticket ID *",lis_ticket1, index=None,
                            placeholder="Select option...",key='delivery')
                    with col2:
                        date_del = st.date_input("Date of Delivery *",value=datetime.today().date())

                    list_delivery = [
                        "Report", "Email Reply", "Dashboard", "New Data Points", "Peer Learning Facilitation", "TA Meeting", "Other"
                    ]

                    type_delivery = st.selectbox(
                        "Type of Delivery *",
                        list_delivery,
                        index=None,
                        placeholder="Select option..."
                    )

                    # If "Other" is selected, show a text input for custom value
                    if type_delivery == "Other":
                        type_delivery_other = st.text_input("Please specify the Type of Delivery *")
                        if type_delivery_other:
                            type_delivery = type_delivery_other 
                    delivery_description = st.text_area("Short Summary *", placeholder='Enter text', height=150,key='delivery_description') 
                    document_del = st.file_uploader(
                        "Upload any files or attachments that are relevant to this delivery.",accept_multiple_files=True
                    )

                    # Submit button
                    st.markdown("""
                        <style>
                        .stButton > button {
                            width: 100%;
                            background-color: #cdb4db;
                            color: black;
                            font-family: Arial, "Segoe UI", sans-serif;
                            font-weight: 600;
                            border-radius: 8px;
                            padding: 0.6em;
                            margin-top: 1em;
                        }
                        </style>
                    """, unsafe_allow_html=True)

                    # Submit logic
                    if st.button("Submit",key='delivery_submit'):
                        errors = []
                        drive_links_del = ""
                        # Required field checks
                        if not ticket_id_del: errors.append("Ticket ID is required.")
                        if not date_del: errors.append("Date of delivery is required.")
                        if not type_delivery: errors.append("Type of delivery is required.")
                        if not delivery_description: errors.append("Short summary is required.")

                        # Show warnings or success
                        if errors:
                            for error in errors:
                                st.warning(error)
                        else:
                            # Only upload files if all validation passes
                            if document_del:
                                try:
                                    folder_id_del = "1gXfWxys2cxd67YDk8zKPmG_mLGID4qL2" 
                                    links_del = []
                                    for file in document_del:
                                        # Rename file as: GU0001_filename.pdf
                                        renamed_filename = f"{ticket_id_del}_{file.name}"
                                        link = upload_file_to_drive(
                                            file=file,
                                            filename=renamed_filename,
                                            folder_id=folder_id_del,
                                            creds_dict=st.secrets["gcp_service_account"]
                                        )
                                        links_del.append(link)
                                    drive_links_del = ", ".join(links_del)
                                    st.success("File(s) uploaded to Google Drive.")    
                                except Exception as e:
                                    st.error(f"Error uploading file(s) to Google Drive: {str(e)}")

                            new_row_del = {
                                'Ticket ID': ticket_id_del,
                                "Date of Delivery": date_del.strftime("%Y-%m-%d"),  # Convert to string
                                "Type of Delivery": type_delivery,
                                "Short Summary": delivery_description,
                                "Document": drive_links_del,
                                "Submitted By": coordinator_name,
                                "Submission Date": datetime.today().strftime("%Y-%m-%d %H:%M")
                            }
                            new_data_del = pd.DataFrame([new_row_del])

                            try:
                                # Append new data to Google Sheet
                                updated_sheet2 = pd.concat([df_del, new_data_del], ignore_index=True)
                                updated_sheet2= updated_sheet2.applymap(
                                    lambda x: x.strftime("%Y-%m-%d") if isinstance(x, (datetime, pd.Timestamp)) else x
                                )
                                # Replace NaN with empty strings to ensure JSON compatibility
                                updated_sheet2 = updated_sheet2.fillna("")
                                spreadsheet3 = client.open('Example_TA_Request')
                                worksheet3 = spreadsheet3.worksheet('Delivery')
                                worksheet3.update([updated_sheet2.columns.values.tolist()] + updated_sheet2.values.tolist())

                                # Clear cache to refresh data
                                st.cache_data.clear()
                                
                                st.success("✅ Submission successful!")
                                time.sleep(2)
                                st.rerun()

                            except Exception as e:
                                st.error(f"Error updating Google Sheets: {str(e)}")

                st.markdown("<hr style='margin:2em 0; border:1px solid #dee2e6;'>", unsafe_allow_html=True)

                with st.expander("📦 **CHECK INTERACTION & DELIVERY PATTERNS**"):
                    st.markdown("""
                        <div style='background: #e3e8fa; border-radius: 18px; box-shadow: 0 4px 18px rgba(26,35,126,0.10); padding: 2em 1.5em 1.5em 1.5em; margin-bottom: 2em; margin-top: 1em;'>
                            <div style='color: #1a237e; font-family: "Segoe UI", "Arial", sans-serif; font-weight: 800; font-size: 1.5em; margin-bottom: 0.4em;'>📦 Explore TA Request Activity</div>
                            <div style='color: #333; font-size: 1.13em; margin-bottom: 0.7em;'>
                                Visualize and analyze the communication and delivery patterns for all Technical Assistance requests. Use the charts and filters below to spot trends, monitor engagement, and drill down into specific Ticket IDs for detailed activity logs.
                            </div>
                        </div>
                    """, unsafe_allow_html=True)
                    # Fetch data from Google Sheets     
                    try:
                        df = load_main_sheet()  # Use cached function
                    except Exception as e:
                        st.error(f"Error fetching data from Google Sheets: {str(e)}")

                    try:
                        df_int = load_interaction_sheet()  # Use cached function
                    except Exception as e:
                        st.error(f"Error fetching data from Google Sheets: {str(e)}")

                    try:
                        df_del = load_delivery_sheet()  # Use cached function
                    except Exception as e:
                        st.error(f"Error fetching data from Google Sheets: {str(e)}")
                    
                    num_interaction = df_int.shape[0]
                    num_delivery = df_del.shape[0]

                    col1, col2 = st.columns(2)
                    col1.metric("🟡 # of Interactions", num_interaction)
                    col2.metric("🟡 # of Deliveries", num_delivery)

                    # Group by Ticket ID and Type of Interaction, count occurrences
                    interaction_counts = df_int.groupby(['Ticket ID', 'Type of Interaction']).size().reset_index(name='Count')
                    delivery_counts = df_del.groupby(['Ticket ID', 'Type of Delivery']).size().reset_index(name='Count')


                    st.markdown("##### 🟡 Top 10 with most Interactions by Interaction Type")
                    if not interaction_counts.empty:
                        pie1 = alt.Chart(interaction_counts).mark_bar().encode(
                            y=alt.Y('Ticket ID:N', sort='-x', title='Ticket ID'),
                            x=alt.X('Count:Q', title='Number of Interactions'),
                            color=alt.Color('Type of Interaction:N', title='Interaction Type'),
                            tooltip=['Ticket ID', 'Type of Interaction', 'Count']
                        ).properties(
                            width=600,
                            height=400
                        )
                        st.altair_chart(pie1, use_container_width=True)
                    else:
                        st.info("No any interaction to show.")

                    st.markdown("##### 🟡 Top 10 with most Deliveries by Delivery Type")
                    if not delivery_counts.empty:
                        pie1 = alt.Chart(delivery_counts).mark_bar().encode(
                            y=alt.Y('Ticket ID:N', sort='-x', title='Ticket ID'),
                            x=alt.X('Count:Q', title='Number of Deliveries'),
                            color=alt.Color('Type of Delivery:N', title='Delivery Type'),
                            tooltip=['Ticket ID', 'Type of Delivery', 'Count']
                        ).properties(
                            width=600,
                            height=400
                        )
                        st.altair_chart(pie1, use_container_width=True)
                    else:
                        st.info("No any delivery to show.")


                    unique_id = sorted(set(df_int["Ticket ID"].unique().tolist() + df_del["Ticket ID"].unique().tolist()))
                    selected_ticket_id = st.selectbox("Select a Ticket ID", unique_id, placeholder="Select option...")
                    # Filter records
                    interactions_for_ticket = df_int[df_int["Ticket ID"] == selected_ticket_id]
                    deliveries_for_ticket = df_del[df_del["Ticket ID"] == selected_ticket_id]
                    num_interaction_ticket = interactions_for_ticket.shape[0]
                    num_delivery_ticket = deliveries_for_ticket.shape[0]
                    st.markdown(f"##### 🟡 Ticket ID: {selected_ticket_id}")
                    st.markdown(f"🟡 # of Interactions: {num_interaction_ticket if num_interaction_ticket > 0 else 0}")
                    if num_interaction_ticket > 0:
                        st.dataframe(interactions_for_ticket)
                    else:
                        st.info("No interaction records found for this Ticket ID.")
                    st.markdown(f"🟡 # of Deliveries: {num_delivery_ticket if num_delivery_ticket > 0 else 0}")
                    if num_delivery_ticket > 0:
                        st.dataframe(deliveries_for_ticket)
                    else:
                        st.info("No delivery records found for this Ticket ID.")
                  


                st.markdown("<hr style='margin:2em 0; border:1px solid #dee2e6;'>", unsafe_allow_html=True)


            elif st.session_state.role == "Assignee/Staff":
                # Add staff content here
                user_info = USERS.get(st.session_state.user_email)
                staff_name = user_info["Assignee/Staff"]["name"] if user_info and "Assignee/Staff" in user_info else None
                st.markdown(
                    """
                    <div style='
                        display: flex;
                        flex-direction: column;
                        align-items: center;
                        justify-content: center;
                        background: #f8f9fa;
                        padding: 2em 0 1em 0;
                        border-radius: 18px;
                        box-shadow: 0 4px 24px rgba(0,0,0,0.07);
                        margin-bottom: 2em;
                    '>
                        <img src='https://raw.githubusercontent.com/JiaqinWu/HRSA64_Dash/main/Georgetown_logo_blueRGB.png' width='200' style='margin-bottom: 1em;'/>
                        <h1 style='
                            color: #1a237e;
                            font-family: "Segoe UI", "Arial", sans-serif;
                            font-weight: 700;
                            margin: 0;
                            font-size: 2.2em;
                            text-align: center;
                        '>👷 Staff Dashboard</h1>
                    </div>
                    """,
                    unsafe_allow_html=True
                )
                #st.header("👷 Staff Dashboard")
                # Personalized greeting
                if staff_name:
                    st.markdown(f"""
                    <div style='                      
                    background: #f8f9fa;                        
                    border-radius: 12px;                        
                    box-shadow: 0 2px 8px rgba(0,0,0,0.04);                        
                    padding: 1.2em 1em 1em 1em;                        
                    margin-bottom: 1.5em;                        
                    text-align: center;                        
                    font-family: Arial, "Segoe UI", sans-serif;                    
                    '>
                        <span style='                           
                        font-size: 1.15em;
                        font-weight: 700;
                        color: #1a237e;
                        letter-spacing: 0.5px;'>
                            👋 Welcome, {staff_name}!
                        </span>
                    </div>
                    """, unsafe_allow_html=True)


                # Filter requests assigned to current staff and In Progress
                staff_df = df[(df["Assigned Coach"] == staff_name) & (df["Status"] == "In Progress")].copy()
                com_df = df[(df["Assigned Coach"] == staff_name) & (df["Status"] == "Completed")].copy()


                # Ensure date columns are datetime
                staff_df["Targeted Due Date"] = pd.to_datetime(staff_df["Targeted Due Date"], errors="coerce")
                staff_df["Assigned Date"] = pd.to_datetime(staff_df["Assigned Date"], errors="coerce")

                # --- Top Summary Cards
                col1, col2 = st.columns(2)
                col3, col4 = st.columns(2)
                # 1. Total In Progress
                total_in_progress = staff_df['Ticket ID'].nunique()
                total_complete = com_df['Ticket ID'].nunique()

                # 2. Newly Assigned: within last 3 days
                recent_cutoff = datetime.today() - timedelta(days=3)
                newly_assigned = staff_df[staff_df["Assigned Date"] >= recent_cutoff]['Ticket ID'].nunique()

                # 3. Due within 1 month
                due_soon_cutoff = datetime.today() + timedelta(days=30)
                due_soon = staff_df[staff_df["Targeted Due Date"] <= due_soon_cutoff]['Ticket ID'].nunique()

                col1.metric("🟡 In Progress", total_in_progress)
                col2.metric("✅ Completed", total_complete)
                col3.metric("🆕 Newly Assigned (Last 3 days)", newly_assigned)
                col4.metric("📅 Due Within 1 Month", due_soon)

                style_metric_cards(border_left_color="#DBF227")


                # --- Section 2: Filter, Sort, Comment
                with st.expander("🚧 **IN-PROGRESS REQUESTS**"):
                    st.markdown("#### 🚧 In-progress Requests")

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
                        st.markdown("##### 🔍 Filter Options")

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
                            "Ticket ID","Jurisdiction", "Organization", "Name", "Title/Position", "Email Address", "Phone Number",
                            "Focus Area", "TA Type", "Assigned Date", "Targeted Due Date","Expected Duration (Days)","Priority", "Assigned Coach", "TA Description",
                            "Document","Coordinator Comment History", "Staff Comment History", "Transfer History"
                        ]].sort_values(by="Expected Duration (Days)").reset_index(drop=True))

                        # Select request by index (row number in submitted_requests)
                        request_indices2 = filtered_df2.index.tolist()
                        selected_request_index1 = st.selectbox(
                            "Select a request to comment",
                            options=request_indices2,
                            format_func=lambda idx: f"{filtered_df2.at[idx, 'Ticket ID']} | {filtered_df2.at[idx, 'Name']} | {filtered_df2.at[idx, 'Jurisdiction']}",
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
                                # Keep latest comment in main field
                                updated_df.loc[global_index, "Staff Comment"] = comment_text
                                # Append to history with timestamp and author
                                ts = datetime.today().strftime("%Y-%m-%d %H:%M")
                                author = staff_name or "Staff"
                                entry = f"{ts} | {author}: {comment_text}" if comment_text else ""
                                if entry:
                                    existing = str(updated_df.loc[global_index, "Staff Comment History"]).strip()
                                    if existing and existing.lower() != "nan":
                                        updated_df.loc[global_index, "Staff Comment History"] = existing + "\n" + entry
                                    else:
                                        updated_df.loc[global_index, "Staff Comment History"] = entry
                                updated_df = updated_df.applymap(
                                    lambda x: x.strftime("%Y-%m-%d") if isinstance(x, (pd.Timestamp, datetime)) and not pd.isna(x) else x
                                )
                                updated_df = updated_df.fillna("") 

                                # Push to Google Sheets
                                spreadsheet1 = client.open('Example_TA_Request')
                                worksheet1 = spreadsheet1.worksheet('Main')
                                worksheet1.update([updated_df.columns.values.tolist()] + updated_df.values.tolist())

                                st.cache_data.clear()

                                st.success("💬 Comment saved successfully!.")
                                time.sleep(2)
                                st.rerun()

                            except Exception as e:
                                st.error(f"Error saving comment: {str(e)}")
                st.markdown("<hr style='margin:2em 0; border:1px solid #dee2e6;'>", unsafe_allow_html=True)

                with st.expander("🗒️ **SUBMIT INTERACTION LOG**"):
                    st.markdown("""
                        <div style='background: #f0f4ff; border-radius: 16px; box-shadow: 0 2px 8px rgba(26,35,126,0.08); padding: 1.5em 1em 1em 1em; margin-bottom: 2em; margin-top: 1em;'>
                            <div style='color: #1a237e; font-family: "Segoe UI", "Arial", sans-serif; font-weight: 700; font-size: 1.4em; margin-bottom: 0.3em;'>🗒️ Submit a New Interaction Log Form</div>
                            <div style='color: #333; font-size: 1.08em; margin-bottom: 0.8em;'>
                                Log a new interaction with a jurisdiction. Fill out the form below to record emails, meetings, or other communications related to a TA request.
                            </div>
                        </div>
                    """, unsafe_allow_html=True)
                    lis_ticket = ["No Ticket ID"] + sorted([tid for tid in df["Ticket ID"].dropna().astype(str).unique().tolist()])

                    # Interaction Log form
                    col1, col2 = st.columns(2)
                    with col1:
                        ticket_id_int = st.selectbox("Ticket ID *", lis_ticket, index=None,
                            placeholder="Select option...", key='interaction1')
                    with col2:
                        date_int = st.date_input("Date of Interaction *",value=datetime.today().date())
                    
                                        
                    # If No Ticket ID, ask for Jurisdiction
                    jurisdiction_for_no_ticket1 = None
                    if ticket_id_int == "No Ticket ID":
                        jurisdiction_for_no_ticket1 = st.selectbox(
                            "Jurisdiction *",
                            lis_location,
                            index=None,
                            placeholder="Select option...",
                            key='juris_interaction1'
                        )

                    list_interaction = [
                        "Email", "Phone Call", "In-Person Meeting", "Online Meeting", "Other"
                    ]

                    type_interaction = st.selectbox(
                        "Type of Interaction *",
                        list_interaction,
                        index=None,
                        placeholder="Select option..."
                    )

                    # If "Other" is selected, show a text input for custom value
                    if type_interaction == "Other":
                        type_interaction_other = st.text_input("Please specify the Type of Interaction *")
                        if type_interaction_other:
                            type_interaction = type_interaction_other 
                    interaction_description = st.text_area("Short Summary *", placeholder='Enter text', height=150,key='interaction_description1') 

                    document_int = st.file_uploader(
                        "Upload any files or attachments that are relevant to this interaction.",accept_multiple_files=True
                    )

                    # Submit button
                    st.markdown("""
                        <style>
                        .stButton > button {
                            width: 100%;
                            background-color: #cdb4db;
                            color: black;
                            font-family: Arial, "Segoe UI", sans-serif;
                            font-weight: 600;
                            border-radius: 8px;
                            padding: 0.6em;
                            margin-top: 1em;
                        }
                        </style>
                    """, unsafe_allow_html=True)

                    try:
                        df = load_main_sheet()  # Use cached function
                    except Exception as e:
                        st.error(f"Error fetching data from Google Sheets: {str(e)}")

                    try:
                        df_int = load_interaction_sheet()  # Use cached function
                    except Exception as e:
                        st.error(f"Error fetching data from Google Sheets: {str(e)}")

                    try:
                        df_del = load_delivery_sheet()  # Use cached function
                    except Exception as e:
                        st.error(f"Error fetching data from Google Sheets: {str(e)}")

                    # Submit logic
                    if st.button("Submit",key='interaction_submit1'):
                        errors = []
                        drive_links_int = ""  # Initialize here
                        # Required field checks
                        if not ticket_id_int: errors.append("Ticket ID is required.")
                        if ticket_id_int == "No Ticket ID" and not jurisdiction_for_no_ticket1:
                            errors.append("Jurisdiction is required when Ticket ID is not provided.")
                        if not date_int: errors.append("Date of interaction is required.")
                        if not type_interaction: errors.append("Type of interaction is required.")
                        if not interaction_description: errors.append("Short summary is required.")

                        # Show warnings or success
                        if errors:
                            for error in errors:
                                st.warning(error)
                        else:
                            # Only upload files if all validation passes
                            if document_int:
                                try:
                                    folder_id_int = "19-Sm8W151tg1zyDN0Nh14DUvOVUieqq7" 
                                    links_int = []
                                    for file in document_int:
                                        # Rename file as: GU0001_filename.pdf
                                        renamed_filename = f"{ticket_id_int}_{file.name}"
                                        link = upload_file_to_drive(
                                            file=file,
                                            filename=renamed_filename,
                                            folder_id=folder_id_int,
                                            creds_dict=st.secrets["gcp_service_account"]
                                        )
                                        links_int.append(link)
                                    drive_links_int = ", ".join(links_int)
                                    st.success("File(s) uploaded to Google Drive.")    
                                except Exception as e:
                                    st.error(f"Error uploading file(s) to Google Drive: {str(e)}")

                            new_row_int = {
                                'Ticket ID': ticket_id_int,
                                "Date of Interaction": date_int.strftime("%Y-%m-%d"),  # Convert to string
                                "Type of Interaction": type_interaction,
                                "Short Summary": interaction_description,
                                "Document": drive_links_int,
                                "Jurisdiction": (lambda: (
                                    str(df.loc[df["Ticket ID"].astype(str) == str(ticket_id_int), "Jurisdiction"].iloc[0])
                                    if ticket_id_int != "No Ticket ID" and not df.loc[df["Ticket ID"].astype(str) == str(ticket_id_int), "Jurisdiction"].empty
                                    else (jurisdiction_for_no_ticket1 or "")
                                ))(),
                                "Submitted By": staff_name,
                                "Submission Date": datetime.today().strftime("%Y-%m-%d %H:%M")
                            }
                            new_data_int = pd.DataFrame([new_row_int])

                            try:
                                # Append new data to Google Sheet
                                updated_sheet2 = pd.concat([df_int, new_data_int], ignore_index=True)
                                updated_sheet2= updated_sheet2.applymap(
                                    lambda x: x.strftime("%Y-%m-%d") if isinstance(x, (datetime, pd.Timestamp)) else x
                                )
                                # Replace NaN with empty strings to ensure JSON compatibility
                                updated_sheet2 = updated_sheet2.fillna("")
                                
                                # Get the worksheet first
                                spreadsheet3 = client.open('Example_TA_Request')
                                worksheet3 = spreadsheet3.worksheet('Interaction')
                                worksheet3.update([updated_sheet2.columns.values.tolist()] + updated_sheet2.values.tolist())

                                # Clear cache to refresh data
                                st.cache_data.clear()
                                
                                st.success("✅ Submission successful!")
                                time.sleep(2)
                                st.rerun()

                            except Exception as e:
                                st.error(f"Error updating Google Sheets: {str(e)}")

                st.markdown("<hr style='margin:2em 0; border:1px solid #dee2e6;'>", unsafe_allow_html=True)

                with st.expander("📦 **SUBMIT DELIVERY FORM**"):
                    st.markdown("""
                        <div style='background: #f0f4ff; border-radius: 16px; box-shadow: 0 2px 8px rgba(26,35,126,0.08); padding: 1.5em 1em 1em 1em; margin-bottom: 2em; margin-top: 1em;'>
                            <div style='color: #1a237e; font-family: "Segoe UI", "Arial", sans-serif; font-weight: 700; font-size: 1.4em; margin-bottom: 0.3em;'>📦 Submit a New Delivery Form</div>
                            <div style='color: #333; font-size: 1.08em; margin-bottom: 0.8em;'>
                                Record a new delivery (e.g., report, dashboard, data) for a TA request. Use the form below to upload files and provide a summary of the delivery.
                            </div>
                        </div>
                    """, unsafe_allow_html=True)
                    lis_ticket1 = df["Ticket ID"].unique().tolist()

                    # Interaction Log form
                    col1, col2 = st.columns(2)
                    with col1:
                        ticket_id_del = st.selectbox("Ticket ID *",lis_ticket1, index=None,
                            placeholder="Select option...",key='delivery1')
                    with col2:
                        date_del = st.date_input("Date of Delivery *",value=datetime.today().date())

                    list_delivery = [
                        "Report", "Email Reply", "Dashboard", "New Data Points","Peer Learning Facilitation", "TA Meeting", "Other"
                    ]

                    type_delivery = st.selectbox(
                        "Type of Delivery *",
                        list_delivery,
                        index=None,
                        placeholder="Select option..."
                    )

                    # If "Other" is selected, show a text input for custom value
                    if type_delivery == "Other":
                        type_delivery_other = st.text_input("Please specify the Type of Delivery *")
                        if type_delivery_other:
                            type_delivery = type_delivery_other 
                    delivery_description = st.text_area("Short Summary *", placeholder='Enter text', height=150,key='delivery_description1') 
                    document_del = st.file_uploader(
                        "Upload any files or attachments that are relevant to this delivery.",accept_multiple_files=True
                    )

                    # Submit button
                    st.markdown("""
                        <style>
                        .stButton > button {
                            width: 100%;
                            background-color: #cdb4db;
                            color: black;
                            font-family: Arial, "Segoe UI", sans-serif;
                            font-weight: 600;
                            border-radius: 8px;
                            padding: 0.6em;
                            margin-top: 1em;
                        }
                        </style>
                    """, unsafe_allow_html=True)

                    # Submit logic
                    if st.button("Submit",key='delivery_submit1'):
                        errors = []
                        drive_links_del = ""  # Ensure always defined
                        # Required field checks
                        if not ticket_id_del: errors.append("Ticket ID is required.")
                        if not date_del: errors.append("Date of delivery is required.")
                        if not type_delivery: errors.append("Type of delivery is required.")
                        if not delivery_description: errors.append("Short summary is required.")

                        # Show warnings or success
                        if errors:
                            for error in errors:
                                st.warning(error)
                        else:
                            # Only upload files if all validation passes
                            if document_del:
                                try:
                                    folder_id_del = "1gXfWxys2cxd67YDk8zKPmG_mLGID4qL2" 
                                    links_del = []
                                    for file in document_del:
                                        # Rename file as: GU0001_filename.pdf
                                        renamed_filename = f"{ticket_id_del}_{file.name}"
                                        link = upload_file_to_drive(
                                            file=file,
                                            filename=renamed_filename,
                                            folder_id=folder_id_del,
                                            creds_dict=st.secrets["gcp_service_account"]
                                        )
                                        links_del.append(link)
                                    drive_links_del = ", ".join(links_del)
                                    st.success("File(s) uploaded to Google Drive.")    
                                except Exception as e:
                                    st.error(f"Error uploading file(s) to Google Drive: {str(e)}")

                            new_row_del = {
                                'Ticket ID': ticket_id_del,
                                "Date of Delivery": date_del.strftime("%Y-%m-%d"),  # Convert to string
                                "Type of Delivery": type_delivery,
                                "Short Summary": delivery_description,
                                "Document": drive_links_del,
                                "Submitted By": staff_name,
                                "Submission Date": datetime.today().strftime("%Y-%m-%d %H:%M")
                            }
                            new_data_del = pd.DataFrame([new_row_del])

                            try:
                                # Append new data to Google Sheet
                                updated_sheet2 = pd.concat([df_del, new_data_del], ignore_index=True)
                                updated_sheet2= updated_sheet2.applymap(
                                    lambda x: x.strftime("%Y-%m-%d") if isinstance(x, (datetime, pd.Timestamp)) else x
                                )
                                # Replace NaN with empty strings to ensure JSON compatibility
                                updated_sheet2 = updated_sheet2.fillna("")
                                spreadsheet3 = client.open('Example_TA_Request')
                                worksheet3 = spreadsheet3.worksheet('Delivery')
                                worksheet3.update([updated_sheet2.columns.values.tolist()] + updated_sheet2.values.tolist())

                                # Clear cache to refresh data
                                st.cache_data.clear()
                                
                                st.success("✅ Submission successful!")
                                time.sleep(2)
                                st.rerun()

                            except Exception as e:
                                st.error(f"Error updating Google Sheets: {str(e)}")

                st.markdown("<hr style='margin:2em 0; border:1px solid #dee2e6;'>", unsafe_allow_html=True)

                # --- Section 1: Mark as Completed
                with st.expander("✅ **MARK REQUESTS AS COMPLETED**"):
                    st.markdown("#### ✅ Mark Requests as Completed")
                    if staff_df.empty:
                        st.info("No requests currently in progress to mark as completed.")
                    else:
                        # Ensure datetime before using .dt
                        staff_df["Assigned Date"] = pd.to_datetime(staff_df["Assigned Date"], errors="coerce")
                        staff_df["Targeted Due Date"] = pd.to_datetime(staff_df["Targeted Due Date"], errors="coerce")

                        # Format dates
                        staff_df["Assigned Date"] = staff_df["Assigned Date"].dt.strftime("%Y-%m-%d")
                        staff_df["Targeted Due Date"] = staff_df["Targeted Due Date"].dt.strftime("%Y-%m-%d")

                        # Display clean table (exclude PriorityOrder column)
                        st.dataframe(staff_df[[
                            "Ticket ID","Jurisdiction", "Organization", "Name", "Title/Position", "Email Address", "Phone Number",
                            "Focus Area", "TA Type", "Assigned Date", "Targeted Due Date", "Priority", "TA Description","Document","Coordinator Comment History"
                        ]].reset_index(drop=True))

                        # Select request by index (row number in submitted_requests)
                        request_indices = staff_df.index.tolist()
                        selected_request_index = st.selectbox(
                            "Select a request to marked as completed",
                            options=request_indices,
                            format_func=lambda idx: f"{staff_df.at[idx, 'Ticket ID']} | {staff_df.at[idx, 'Name']} | {staff_df.at[idx, 'Jurisdiction']}",
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

                                updated_df = updated_df.applymap(
                                    lambda x: x.strftime("%Y-%m-%d") if isinstance(x, (pd.Timestamp, datetime)) and not pd.isna(x) else x
                                )
                                updated_df = updated_df.fillna("") 
                                spreadsheet1 = client.open('Example_TA_Request')
                                worksheet1 = spreadsheet1.worksheet('Main')

                                # Push to Google Sheet
                                worksheet1.update([updated_df.columns.values.tolist()] + updated_df.values.tolist())

                                # Clear cache to refresh data
                                st.cache_data.clear()
                                
                                st.success("✅ Request marked as completed.")
                                time.sleep(2)
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

                st.markdown("<hr style='margin:2em 0; border:1px solid #dee2e6;'>", unsafe_allow_html=True)

