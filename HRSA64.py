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


# Example usage: Fetch data from Google Sheets
try:
    spreadsheet1 = client.open('Example_TA_Request')
    worksheet1 = spreadsheet1.worksheet('Main')
    df = pd.DataFrame(worksheet1.get_all_records())
except Exception as e:
    st.error(f"Error fetching data from Google Sheets: {str(e)}")

df['Submit Date'] = pd.to_datetime(df['Submit Date'], errors='coerce')
df["Phone Number"] = df["Phone Number"].astype(str)

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
        "Coordinator": {"password": "Qin8851216!", "name": "Jiaqin Wu"},
        "Assignee/Staff": {"password": "Qin8851216!", "name": "Jiaqin Wu"}
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
    st.set_page_config(
        page_title="GU TAP System",
        page_icon="https://raw.githubusercontent.com/JiaqinWu/HRSA64_Dash/main/Georgetown_logo_blueRGB.png", 
        layout="centered"
    ) 
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
                        folder_id = "1Q9dMMdyfEGWFVv2_CbHbJVMHXOST3OYf" 
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
                    worksheet1.update([updated_sheet.columns.values.tolist()] + updated_sheet.values.tolist())
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
                    for key in list(st.session_state.keys()):
                        del st.session_state[key]
                    time.sleep(5)
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
                total_request = df.shape[0]
                inprogress_request = df[df['Status'] == 'In Progress'].shape[0]
                completed_request = df[df['Status'] == 'Completed'].shape[0]

                col1.metric(label="# of Total Requests", value= millify(total_request, precision=2))
                col2.metric(label="# of In-Progress Requests", value= millify(inprogress_request, precision=2))
                col3.metric(label="# of Completed Requests", value= millify(completed_request, precision=2))
                style_metric_cards(border_left_color="#DBF227")

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
                        '>📊 Monitoring In-Progress TA Requests</h2>
                        <p style='
                            color: #333;
                            font-size: 1em;
                            margin-bottom: 0.8em;
                        '>
                            This section provides an overview and monitoring tools for all Technical Assistance (TA) requests that are currently in progress. Use the charts and filters below to track assignments, due dates, and staff workload.
                        </p>
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
                "Martine Etienne-Mesubi", "Seble Kassaye", "Weijun Yu", "Jiaqin Wu", "Zelalem Temesgen", "Carlos Rodriguez-Diaz"]

                staff_list_sorted = sorted(staff_list, key=lambda x: x.split()[0])

                selected_staff = st.selectbox("Select a staff to view their requests", staff_list_sorted, index=None,
                        placeholder="Select option...")

                today = datetime.today()
                last_month = today - timedelta(days=30)
                staff_dff = df[df["Assigned Coach"] == selected_staff].copy()
                in_progress_count = staff_dff[staff_dff["Status"] == "In Progress"].shape[0]
                due_soon_count = staff_dff[
                    (staff_dff["Status"] == "In Progress") & 
                    (staff_dff["Targeted Due Date"] <= next_month)
                ].shape[0]
                completed_recently = staff_dff[
                    (staff_dff["Status"] == "Completed") & 
                    (staff_dff["Submit Date"] >= last_month)
                ].shape[0]

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
                st.markdown("")
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
                        '>📋 Assign TA Requests</h2>
                        <p style='
                            color: #333;
                            font-size: 1em;
                            margin-bottom: 0.8em;
                        '>
                            This section lists all TA requests that have not yet been assigned to a coach. Review the details and assign a staff member to start the TA process.
                        </p>
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

                            # Push to Google Sheet
                            worksheet1.update([updated_df.columns.values.tolist()] + updated_df.values.tolist())

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

                    st.markdown("<hr style='margin:2em 0; border:1px solid #dee2e6;'>", unsafe_allow_html=True)
                    st.markdown("")

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
                            "Assigned Coach", "TA Description","Document","Coordinator Comment"
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
                                updated_df.loc[selected_row_global_index, "Coordinator Comment"] = comment_input
                                updated_df = updated_df.applymap(
                                    lambda x: x.strftime("%Y-%m-%d") if isinstance(x, (pd.Timestamp, datetime)) and not pd.isna(x) else x
                                )
                                updated_df = updated_df.fillna("") 

                                # Push to Google Sheets
                                worksheet1.update([updated_df.columns.values.tolist()] + updated_df.values.tolist())

                                st.success("💬 Comment saved and synced with Google Sheets.")
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
                            'Actual Duration (Days)', "Coordinator Comment", "Staff Comment"
                        ]].reset_index(drop=True))

                        st.markdown("<hr style='margin:2em 0; border:1px solid #dee2e6;'>", unsafe_allow_html=True)
                        st.markdown("")

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
                total_in_progress = staff_df.shape[0]
                total_complete = com_df.shape[0]

                # 2. Newly Assigned: within last 3 days
                recent_cutoff = datetime.today() - timedelta(days=3)
                newly_assigned = staff_df[staff_df["Assigned Date"] >= recent_cutoff].shape[0]

                # 3. Due within 1 month
                due_soon_cutoff = datetime.today() + timedelta(days=30)
                due_soon = staff_df[staff_df["Targeted Due Date"] <= due_soon_cutoff].shape[0]

                col1.metric("🟡 In Progress", total_in_progress)
                col2.metric("✅ Completed", total_complete)
                col3.metric("🆕 Newly Assigned (Last 3 days)", newly_assigned)
                col4.metric("📅 Due Within 1 Month", due_soon)

                style_metric_cards(border_left_color="#DBF227")

                # --- Section 1: Mark as Completed
                st.markdown("#### ✅ Mark Requests as Completed")
                # Format dates
                staff_df["Assigned Date"] = staff_df["Assigned Date"].dt.strftime("%Y-%m-%d")
                staff_df["Targeted Due Date"] = staff_df["Targeted Due Date"].dt.strftime("%Y-%m-%d")

                # Display clean table (exclude PriorityOrder column)
                st.dataframe(staff_df[[
                    "Ticket ID","Jurisdiction", "Organization", "Name", "Title/Position", "Email Address", "Phone Number",
                    "Focus Area", "TA Type", "Assigned Date", "Targeted Due Date", "Priority", "TA Description","Document","Coordinator Comment"
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

                        # Push to Google Sheet
                        worksheet1.update([updated_df.columns.values.tolist()] + updated_df.values.tolist())

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

                # --- Section 2: Filter, Sort, Comment
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
                        "Document","Coordinator Comment", "Staff Comment"
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
                            updated_df.loc[global_index, "Staff Comment"] = comment_text
                            updated_df = updated_df.applymap(
                                lambda x: x.strftime("%Y-%m-%d") if isinstance(x, (pd.Timestamp, datetime)) and not pd.isna(x) else x
                            )
                            updated_df = updated_df.fillna("") 

                            # Push to Google Sheets
                            worksheet1.update([updated_df.columns.values.tolist()] + updated_df.values.tolist())

                            st.success("💬 Comment saved successfully!.")
                            time.sleep(2)
                            st.rerun()

                        except Exception as e:
                            st.error(f"Error saving comment: {str(e)}")


