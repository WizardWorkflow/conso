import streamlit as st
import pandas as pd
import smtplib
from pymongo import MongoClient
from email.message import EmailMessage
from email_validator import validate_email, EmailNotValidError
import random
import io

# MongoDB client setup (store connection details in Streamlit secrets)
client = MongoClient(st.secrets["MONGO_URI"])
db = client["TestDB"]
users_collection = db["users"]
feedback_collection = db["feedback"]

# Initialize session state for selected sheets
if "selected_sheets" not in st.session_state:
    st.session_state.selected_sheets = {}

# Function to reset selections
def reset_selections():
    st.session_state.selected_sheets.clear()

# -------------- Utility Functions --------------

# Send OTP for email verification
def send_otp(email, otp):
    msg = EmailMessage()
    msg.set_content(f"Your OTP code is {otp}")
    msg['Subject'] = 'Your OTP Code'
    msg['From'] = st.secrets["EMAIL_USER"]
    msg['To'] = email

    try:
        with smtplib.SMTP_SSL('smtp.gmail.com', 465) as smtp:
            smtp.login(st.secrets["EMAIL_USER"], st.secrets["EMAIL_PASS"])
            smtp.send_message(msg)
        st.success(f"OTP sent to {email}")
    except Exception as e:
        st.error(f"Error sending OTP: {str(e)}")

# Authenticate user via MongoDB
def authenticate_user(email, password):
    user = users_collection.find_one({"email": email, "password": password})
    return user is not None

# Register a new user
def register_user(email, password):
    existing_user = users_collection.find_one({"email": email})
    if existing_user:
        st.error("User already exists.")
    else:
        users_collection.insert_one({"email": email, "password": password})
        st.success("Registration successful.")

# Change user password
def change_password(email, old_password, new_password):
    user = users_collection.find_one({"email": email, "password": old_password})
    if user:
        users_collection.update_one({"email": email}, {"$set": {"password": new_password}})
        st.success("Password changed successfully.")
    else:
        st.error("Invalid email or old password.")

# Consolidation logic for combining CSV/Excel files or specific sheets
def consolidate_files(file_list, file_type):
    consolidated_data = pd.DataFrame()

    for uploaded_file in file_list:
        try:
            if file_type == 'csv':
                df = pd.read_csv(uploaded_file)
                df['Filename'] = uploaded_file.name  # Add filename column
                consolidated_data = pd.concat([consolidated_data, df], ignore_index=True)
            elif file_type == 'excel':
                df = pd.read_excel(uploaded_file, engine='openpyxl')
                df['Filename'] = uploaded_file.name  # Add filename column
                consolidated_data = pd.concat([consolidated_data, df], ignore_index=True)
        except ValueError as e:
            st.error(f"Error processing {uploaded_file.name}: {e}")
        except Exception as e:
            st.error(f"Unexpected error with file {uploaded_file.name}: {e}")

    return consolidated_data

# Consolidation logic for sheets
def consolidate_sheets(file_list, selected_sheets):
    consolidated_data = pd.DataFrame()

    for uploaded_file in file_list:
        try:
            if uploaded_file.name in selected_sheets:
                for sheet in selected_sheets[uploaded_file.name]:
                    sheet_df = pd.read_excel(uploaded_file, sheet_name=sheet)
                    sheet_df['Filename'] = uploaded_file.name  # Add filename column
                    sheet_df['Sheet Name'] = sheet  # Add sheet name column
                    consolidated_data = pd.concat([consolidated_data, sheet_df], ignore_index=True)
        except ValueError as e:
            st.error(f"Error processing sheet in {uploaded_file.name}: {e}")
        except Exception as e:
            st.error(f"Unexpected error with file {uploaded_file.name}: {e}")

    return consolidated_data

# -------------- Streamlit Web Interface --------------

# Main UI
st.title('Data Consolidation Tool')
st.write("Please log in and upload your files for consolidation.")

# User Authentication
st.header('User Authentication')

auth_choice = st.selectbox('Select action:', ['Login', 'Register', 'Change Password'])

email = st.text_input('Enter your email:')
password = st.text_input('Enter your password:', type='password')

if auth_choice == 'Register':
    if st.button('Register'):
        register_user(email, password)
elif auth_choice == 'Change Password':
    new_password = st.text_input('Enter new password:', type='password')
    if st.button('Change Password'):
        change_password(email, password, new_password)
else:
    if st.button('Login'):
        if authenticate_user(email, password):
            st.success('User logged in successfully!')
        else:
            st.error('Invalid email or password.')

# Only show file upload section if user is authenticated
if authenticate_user(email, password):

    # 1. Select File Type (Default is Excel)
    st.header('File Consolidation')

    file_type = st.selectbox("Choose the file type to be consolidated:", options=['excel', 'csv'], index=0)  # Default to 'excel'

    # 2. Select Consolidation Type
    consolidation_type = st.selectbox("Select consolidation type:", ["Consolidate data from files", "Consolidate data from sheets"])

    # 3. Upload Files
    uploaded_files = st.file_uploader(f"Upload {file_type.upper()} files", type=["xlsx", "csv"] if file_type == 'csv' else ["xlsx"], accept_multiple_files=True)

    # Display total number of uploaded files
    if uploaded_files:
        st.write(f"Total uploaded files: {len(uploaded_files)}")

        # 4. File Consolidation
        if consolidation_type == "Consolidate data from files":
            consolidated_data = consolidate_files(uploaded_files, file_type)

            if not consolidated_data.empty:
                # Display consolidated data
                st.write("Consolidated Data from Files:")
                st.dataframe(consolidated_data)

                # Allow the user to name the consolidated output file
                output_file_name = st.text_input("Enter name for the consolidated output file (without extension)", value="consolidated")

                # Option to download consolidated file
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                    consolidated_data.to_excel(writer, index=False)
                st.download_button(label="Download Consolidated File", data=output.getvalue(), file_name=f"{output_file_name}.xlsx")
            else:
                st.warning("No data to display after consolidation.")

        # 5. Sheet Selection for Sheet Consolidation
        elif consolidation_type == "Consolidate data from sheets" and file_type == 'excel':
            selected_sheets = st.session_state.selected_sheets  # Maintain state of selected sheets across reruns
            search_term = st.text_input("Search for sheets (optional)")

            # List to store all matching sheets across all files for the search term
            global_matching_sheets = []

            for file in uploaded_files:
                sheet_names = pd.ExcelFile(file).sheet_names  # Load sheet names

                # Show manual sheet selection for each file
                selected_sheets[file.name] = st.multiselect(f"Select sheets in {file.name} manually:", options=sheet_names, default=selected_sheets.get(file.name, []))

                # Apply search query to filter sheets
                filtered_sheets = [sheet for sheet in sheet_names if search_term.lower() in sheet.lower()]

                # Store sheets matching the search term globally (filename - sheet name format)
                for sheet in filtered_sheets:
                    global_matching_sheets.append((file.name, sheet))

            # Automatically select and consolidate sheets matching the search term
            if search_term and global_matching_sheets:
                st.write(f"Automatically consolidating sheets matching '{search_term}' across all files:")
                for filename, sheet_name in global_matching_sheets:
                    if filename not in selected_sheets:
                        selected_sheets[filename] = []
                    selected_sheets[filename].append(sheet_name)

            # Consolidate the selected sheets
            mapped_sheets = {file.name: selected_sheets[file.name] for file in uploaded_files if selected_sheets.get(file.name)}
            consolidated_data = consolidate_sheets(uploaded_files, mapped_sheets)

            if not consolidated_data.empty:
                # Display consolidated data
                st.write("Consolidated Data from Sheets:")
                st.dataframe(consolidated_data)

                # Allow the user to name the consolidated output file
                output_file_name = st.text_input("Enter name for the consolidated output file (without extension)", value="consolidated")

                # Option to download consolidated file
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                    consolidated_data.to_excel(writer, index=False)
                st.download_button(label="Download Consolidated File", data=output.getvalue(), file_name=f"{output_file_name}.xlsx")
            else:
                st.warning("No data to display after consolidation.")

        # Add Reset Button to Clear Selections
        if st.button("Reset Selections"):
            reset_selections()
            st.success("Selections have been reset.")

    # Feedback section after consolidation
    st.write("Please provide feedback for the consolidation process:")
    feedback = st.text_area("Enter your feedback:")
    
    if st.button("Submit Feedback"):
        feedback_collection.insert_one({"email": email, "feedback": feedback})
        st.success("Feedback submitted successfully!")

# Inject Custom CSS for Fixed Footer
st.markdown("""
    <style>
    /* Fixed footer with the signature and contact info */
    .footer {
        position: fixed;
        left: 0;
        bottom: 0;
        width: 100%;
        background-color: white;
        color: black;
        text-align: center;
        padding: 10px;
        font-size: small;
        border-top: 1px solid #eaeaea;
        z-index: 9999;
    }

    /* Center the signature */
    .footer .signature {
        display: inline-block;
        margin: 0 auto;
    }

    /* Align the phone number to the right */
    .footer .contact-info {
        position: absolute;
        right: 20px;
        bottom: 10px;
        font-size: small;
        color: black;
    }
    </style>

    <div class="footer">
        <div class="signature">By Ansh Gandhi</div>
        <div class="contact-info">+91 75888 34433</div>
    </div>
""", unsafe_allow_html=True)
