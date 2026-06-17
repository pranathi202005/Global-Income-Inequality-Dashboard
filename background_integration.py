import streamlit as st
from openai import OpenAI
from datetime import datetime
import pandas as pd
import os
import re

# Page Configuration
# ✅ Must be first Streamlit command
st.set_page_config(
    page_title="Global Income Inequality Dashboard",
    page_icon="🌍",
    layout="wide",
    initial_sidebar_state="expanded"
)


client = OpenAI(
    api_key=st.secrets["OPENAI_API_KEY"]
)


CREDENTIALS_FILE = 'credentials.xlsx'
CHAT_HISTORY_FILE = "credentials.xlsx"
CHAT_HISTORY_SHEET = "chat_history"

def load_chat_history(username):
    """Load previous chats for a specific user."""
    if os.path.exists(CHAT_HISTORY_FILE):
        try:
            df = pd.read_excel(CHAT_HISTORY_FILE, sheet_name=CHAT_HISTORY_SHEET)
            user_chats = df[df["username"] == username]
            return user_chats.to_dict("records")
        except Exception:
            return []
    return []

# Ensure credentials file exists

def initialize_excel():
    """Initialize Excel file with all sheets if it doesn't exist"""
    if not os.path.exists(CREDENTIALS_FILE):
        # Create DataFrames for all sheets
        user_df = pd.DataFrame(columns=["username", "email", "password"])
        feedback_df = pd.DataFrame(columns=["username", "feedback", "timestamp"])
        profile_df = pd.DataFrame(columns=['username', 'first_name', 'last_name', 'email', 
                                         'contact_number', 'profile_type', 'country', 'last_updated'])
        
        # Save to Excel with multiple sheets
        with pd.ExcelWriter(CREDENTIALS_FILE, engine='openpyxl') as writer:
            user_df.to_excel(writer, sheet_name='user_details', index=False)
            feedback_df.to_excel(writer, sheet_name='feedback', index=False)
            profile_df.to_excel(writer, sheet_name='profiles', index=False)

def load_feedback_data():
    """Load feedback data from Excel file"""
    try:
        df = pd.read_excel(CREDENTIALS_FILE, sheet_name='feedback')
        # Ensure required columns exist
        required_cols = ['username', 'feedback', 'timestamp']
        for col in required_cols:
            if col not in df.columns:
                df[col] = ""
        return df
    except Exception as e:
        # If feedback sheet doesn't exist, return empty DataFrame
        return pd.DataFrame(columns=["username", "feedback", "timestamp"])

def load_user_credentials():
    """Load user credentials from Excel file"""
    try:
        df = pd.read_excel(CREDENTIALS_FILE, sheet_name='user_details')
        # Ensure required columns exist
        required_cols = ['username', 'email', 'password']
        for col in required_cols:
            if col not in df.columns:
                df[col] = ""
        return df
    except:
        return pd.DataFrame(columns=["username", "email", "password"])

def save_user_credentials(df):
    """Save user credentials to Excel file while preserving feedback sheet"""
    try:
        # Load existing feedback
        feedback_df = load_feedback_data()
        
        # Save both sheets
        with pd.ExcelWriter(CREDENTIALS_FILE, engine='openpyxl') as writer:
            df.to_excel(writer, sheet_name='user_details', index=False)
            feedback_df.to_excel(writer, sheet_name='feedback', index=False)
        return True
    except Exception as e:
        st.error(f"Error saving credentials: {e}")
        return False
    
def save_profile_to_excel():
    """Save profile data to Excel file"""
    try:
        # Try to load existing profile data
        try:
            profile_df = pd.read_excel(CREDENTIALS_FILE, sheet_name='profiles')
        except:
            profile_df = pd.DataFrame(columns=['username', 'first_name', 'last_name', 'email', 
                                             'contact_number', 'profile_type', 'country', 'last_updated'])
        
        # Remove existing entry for this user if it exists
        profile_df = profile_df[profile_df['username'] != st.session_state.username]
        
        # Add new profile data
        new_profile = pd.DataFrame([{
            'username': st.session_state.username,
            'first_name': st.session_state.profile_data['first_name'],
            'last_name': st.session_state.profile_data['last_name'],
            'email': st.session_state.profile_data['email'],
            'contact_number': st.session_state.profile_data['contact_number'],
            'profile_type': st.session_state.profile_data['profile_type'],
            'country': st.session_state.profile_data['country'],
            'last_updated': datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        }])
        
        profile_df = pd.concat([profile_df, new_profile], ignore_index=True)
        
        # Load other sheets (try-except for each in case they don't exist)
        try:
            user_df = pd.read_excel(CREDENTIALS_FILE, sheet_name='user_details')
        except:
            user_df = pd.DataFrame(columns=["username", "email", "password"])
            
        try:
            feedback_df = pd.read_excel(CREDENTIALS_FILE, sheet_name='feedback')
        except:
            feedback_df = pd.DataFrame(columns=["username", "feedback", "timestamp"])
        
        # Save all sheets back to Excel
        with pd.ExcelWriter(CREDENTIALS_FILE, engine='openpyxl') as writer:
            user_df.to_excel(writer, sheet_name='user_details', index=False)
            feedback_df.to_excel(writer, sheet_name='feedback', index=False)
            profile_df.to_excel(writer, sheet_name='profiles', index=False)
        
        return True
        
    except Exception as e:
        st.error(f"Error saving profile: {str(e)}")
        return False

def is_valid_email(email):
    """Check if email is valid"""
    pattern = r'^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$'
    return re.match(pattern, email) is not None

# Save feedback
def load_feedback_data():
    """Load feedback data from Excel file"""
    try:
        df = pd.read_excel(CREDENTIALS_FILE, sheet_name='feedback')
        # Ensure required columns exist
        required_cols = ['username', 'feedback', 'timestamp']
        for col in required_cols:
            if col not in df.columns:
                df[col] = ""
        return df
    except Exception as e:
        # If feedback sheet doesn't exist, return empty DataFrame
        return pd.DataFrame(columns=["username", "feedback", "timestamp"])


# Feedback form UI
def show_feedback():
    st.subheader("💬 Send Us Your Feedback")

    if "username" not in st.session_state or not st.session_state.get("logged_in", False):
        st.error("⚠️ Please log in first to submit feedback.")
        return

    feedback_text = st.text_area("Your Feedback:", height=150)

    if st.button("Submit Feedback"):
        if feedback_text.strip():
            save_feedback(st.session_state.username, feedback_text.strip())
            st.success("✅ Thank you! Your feedback has been recorded.")
        else:
            st.warning("⚠️ Please enter some feedback before submitting.")

def save_feedback(username, feedback_text):
    """Save feedback to Excel file"""
    try:
        # Load existing feedback
        feedback_df = load_feedback_data()
        
        # Create new feedback entry
        new_feedback = {
            'username': username,
            'feedback': feedback_text,
            'timestamp': datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        }
        
        # Append new feedback
        feedback_df = pd.concat([feedback_df, pd.DataFrame([new_feedback])], ignore_index=True)
        
        # Load user data to preserve it
        user_df = load_user_credentials()
        
        # Try to load profile data if it exists
        try:
            profile_df = pd.read_excel(CREDENTIALS_FILE, sheet_name='profiles')
        except:
            profile_df = pd.DataFrame(columns=['username', 'first_name', 'last_name', 'email', 
                                             'contact_number', 'profile_type', 'country', 'last_updated'])
        
        # Save all sheets back to Excel
        with pd.ExcelWriter(CREDENTIALS_FILE, engine='openpyxl') as writer:
            user_df.to_excel(writer, sheet_name='user_details', index=False)
            feedback_df.to_excel(writer, sheet_name='feedback', index=False)
            profile_df.to_excel(writer, sheet_name='profiles', index=False)
        
        return True
    except Exception as e:
        st.error(f"Error saving feedback: {e}")
        return False

def show_feedback():
    st.subheader("💬 Send Us Your Feedback")

    if "username" not in st.session_state or not st.session_state.get("logged_in", False):
        st.error("⚠️ Please log in to submit feedback.")
        return

    feedback_text = st.text_area("Your Feedback:", height=150, placeholder="Share your thoughts, suggestions, or issues...")

    if st.button("Submit Feedback"):
        if feedback_text.strip():
            if save_feedback(st.session_state.username, feedback_text.strip()):
                st.success("✅ Feedback saved successfully!")
                #st.rerun()
            else:
                st.error("❌ Failed to save feedback. Please try again.")
        else:
            st.warning("⚠️ Please enter feedback before submitting.")
    


st.markdown(
    """
    <style>
        body {
            background-image: url("https://www.imf.org/-/media/Images/IMF/FANDD/hero/2022/March/Global-Inequalities-Image-PT.ashx");
            background-color: #cccccc;
            background-size: cover;
            background-attachment: fixed;
            background-position: center;
            font-family: 'Segoe UI', sans-serif;
        }
        /* Box around login */
        .login-box {
            background: rgba(0, 0, 0, 0.7);
            padding: 30px;
            border-radius: 15px;
            box-shadow: 0px 6px 18px rgba(0,0,0,0.6);
            max-width: 500px;
            margin: 80px auto;
            color: #black !important;
            font-size: 18px !important; /* Increase overall text size */
        }
        /* Force all text inside forms */
        .login-box * {
            color: #black !important;
            font-size: 18px !important;
        }
        h1, h2, h3, h4, h5 {
            text-align: center;
            color: #FAFDD6 !important;
            font-weight: bold;
            font-size: 26px !important; /* Bigger heading */
        }
        label, span, p {
            color: #EAF2F8 !important;
            font-size: 18px !important;
        }
        /* Style radio buttons */
        div[data-baseweb="radio"] label {
            color: #black !important;
            font-weight: 600;
            font-size: 18px !important;
        }
        /* Style text inputs */
        input {
            border: 1px solid #91ADC8 !important;
            border-radius: 8px !important;
            padding: 10px !important;
            font-size: 18px !important;
            color: #black !important;  
            background: rgba(255,255,255,0.15) !important;
        }
        input::placeholder {
            color: #000000 !important; /* Black placeholder */
            font-size: 18px !important;
        }
        /* Style buttons */
        button[kind="secondary"], button[kind="primary"] {
            background-color: #647FBC !important;
            color: #black !important;
            border-radius: 8px !important;
            font-weight: bold;
            border: none !important;
            font-size: 18px !important;
            padding: 8px 20px !important;
        }
        button[kind="secondary"]:hover, button[kind="primary"]:hover {
            background-color: #3B5998 !important;
            color: #FAFDD6 !important;
        }
        /* Style success/error/info messages */
        .stAlert {
            color: #FFFFFF !important;
            font-weight: 500;
            font-size: 18px !important;
        }
    </style>
    """,
    unsafe_allow_html=True
)


def show_signup():
    st.subheader('🔐 Create Account')
    
    col1, col2 = st.columns(2)
    
    with col1:
        new_username = st.text_input('Username', placeholder="Choose a username", key="signup_username")
    
    with col2:
        new_email = st.text_input('Email Address', placeholder="your.email@example.com", key="signup_email")
    
    new_password = st.text_input('Password', type='password', placeholder="Create a strong password", key="signup_password")
    confirm_password = st.text_input('Confirm Password', type='password', placeholder="Re-enter your password", key="signup_confirm")
    
    signup_btn = st.button('Create Account', use_container_width=True, key="signup_btn")
    
    if signup_btn:
        if not new_username or not new_email or not new_password:
            st.error('❌ Please fill in all fields.')
        elif new_password != confirm_password:
            st.error('❌ Passwords do not match.')
        elif not re.match(r'^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$', new_email):
            st.error('❌ Please enter a valid email address.')
        else:
            credentials_df = load_user_credentials()
            if new_username in credentials_df['username'].values:
                st.error('❌ Username already exists.')
            elif new_email in credentials_df['email'].values:
                st.error('❌ Email address already registered.')
            else:
                new_entry = pd.DataFrame({
                    'username': [new_username], 
                    'email': [new_email], 
                    'password': [new_password]
                })
                credentials_df = pd.concat([credentials_df, new_entry], ignore_index=True)
                if save_user_credentials(credentials_df):
                    st.success('✅ Account created successfully! Please proceed to Login.')
                    st.balloons()
                else:
                    st.error('❌ Failed to create account. Please try again.')

def show_login():
    st.subheader('🔑 Login to Dashboard')
    
    login_option = st.radio("Login with:", ["Username", "Email"], horizontal=True, key="login_option")
    
    if login_option == "Username":
        username = st.text_input('Username', placeholder="Enter your username", key="login_username")
    else:
        username = st.text_input('Email', placeholder="Enter your email address", key="login_email")
    
    password = st.text_input('Password', type='password', placeholder="Enter your password", key="login_password")
    
    col1, col2, col3 = st.columns([1, 2, 1])
    with col2:
        login_btn = st.button('Login', use_container_width=True, key="login_btn")
    
    if login_btn:
        credentials_df = load_user_credentials()
        if login_option == "Username":
            user_match = (credentials_df['username'] == username) & (credentials_df['password'] == password)
        else:
            user_match = (credentials_df['email'] == username) & (credentials_df['password'] == password)
            
        if user_match.any():
            st.session_state.logged_in = True
            user_data = credentials_df[user_match].iloc[0]
            st.session_state.username = user_data['username'] 
            st.session_state.email = user_data['email']
            
            # Initialize profile data first
            st.session_state.profile_data = {
                'first_name': '',
                'last_name': '', 
                'email': st.session_state.email,
                'contact_number': '',
                'profile_type': 'Student',
                'country': '',
                'profile_pic': None
            }
            
            # Then load existing profile data
            load_profile_data()
            
            st.success('✅ Login successful!')
            st.rerun()
        else:
            st.error('❌ Invalid credentials.')


def save_chat_history(username, role, content):
    """Save chat message to Excel under the right user."""
    new_entry = pd.DataFrame([{"username": username, "role": role, "content": content}])

    if os.path.exists(CHAT_HISTORY_FILE):
        try:
            df = pd.read_excel(CHAT_HISTORY_FILE, sheet_name=CHAT_HISTORY_SHEET)
            df = pd.concat([df, new_entry], ignore_index=True)
        except Exception:
            df = new_entry
        with pd.ExcelWriter(CHAT_HISTORY_FILE, engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
            df.to_excel(writer, sheet_name=CHAT_HISTORY_SHEET, index=False)
    else:
        with pd.ExcelWriter(CHAT_HISTORY_FILE, engine="openpyxl") as writer:
            new_entry.to_excel(writer, sheet_name=CHAT_HISTORY_SHEET, index=False)

def show_chatbot():
    st.markdown("## 🤖 AI Chatbot Assistant")

    # Initialize history
    if "chat_history" not in st.session_state:
        st.session_state.chat_history = load_chat_history(st.session_state.username)

    # Toggle history inside chatbot page
    if st.button("📜 Show/Hide Chat History"):
        st.session_state.show_history = not st.session_state.get("show_history", False)

    if st.session_state.get("show_history", False):
        st.markdown("### 📜 Your Previous Conversations")
        if st.session_state.chat_history:
            for chat in st.session_state.chat_history:
                role = "🧑 You" if chat["role"] == "user" else "🤖 Bot"
                st.markdown(f"**{role}:** {chat['content']}")
        else:
            st.info("No previous chat history found.")

    # Chat input
    user_message = st.text_input("Type your message:", key="chat_input")
    if st.button("Send"):
        if user_message.strip():
            # Save user message (in memory + Excel)
            st.session_state.chat_history.append({"role": "user", "content": user_message})
            save_chat_history(st.session_state.username, "user", user_message)

            # Get AI response
            response = client.chat.completions.create(
                model="gpt-4o-mini",
                messages=[{"role": "system", "content": "You are a helpful assistant."}] +
                         [{"role": c["role"], "content": c["content"]} for c in st.session_state.chat_history]
            )
            bot_reply = response.choices[0].message.content

            # Save bot reply (in memory + Excel)
            st.session_state.chat_history.append({"role": "assistant", "content": bot_reply})
            save_chat_history(st.session_state.username, "assistant", bot_reply)

            # Show reply
            st.markdown(f"**🤖 Bot:** {bot_reply}")

def show_home():
    st.markdown("""
        <div style="background: #384959; 
                    padding: 2.5rem; border-radius: 20px; margin-bottom: 2rem; 
                    box-shadow: 0 8px 25px rgba(0, 0, 0, 0.4);">
            <h1 style="color: #BDDDFC; margin: 0; text-align: center;">🌍 Global Income Inequality Dashboard</h1>
            <p style="color: #BDDDFC; font-size: 1.2rem; margin: 0.5rem 0 0 0; text-align: center;">
                Comprehensive Analysis of Global Income Distribution Patterns
            </p>
        </div>
    """, unsafe_allow_html=True)
    
    col1, col2 = st.columns([2, 1])
    
    with col1:
        st.markdown("""
            <div style="background: #384959; 
                        padding: 2rem; border-radius: 15px; margin-bottom: 2rem;
                        border-left: 5px solid #BDDDFC; box-shadow: 0 5px 15px rgba(0, 0, 0, 0.2);">
                <h2 style="color: #BDDDFC; border-bottom: 2px solid #BDDDFC; padding-bottom: 0.5rem;">
                    Welcome to the Dashboard
                </h2>
                <p style="color: #BDDDFC; line-height: 1.7;">
                    This interactive platform provides comprehensive insights into global income inequality trends, 
                    allowing you to explore data visualizations, compare regions, and analyze economic disparities 
                    across different countries and demographics.
                </p>
            </div>
            
            <div style="background: #384959; 
                        padding: 2rem; border-radius: 15px; margin-bottom: 2rem;
                        border-left: 5px solid #BDDDFC; box-shadow: 0 5px 15px rgba(0, 0, 0, 0.2);">
                <h3 style="color: #BDDDFC; border-bottom: 2px solid #BDDDFC; padding-bottom: 0.5rem;">
                    📊 Key Metrics
                </h3>
                <div style="display: grid; grid-template-columns: repeat(2, 1fr); gap: 1rem;">
                    <div style="background: #2a3845; padding: 1rem; border-radius: 10px; text-align: center;">
                        <h4 style="color: #BDDDFC; margin: 0;">Gini Coefficient</h4>
                        <p style="color: #BDDDFC; font-size: 1.5rem; margin: 0.5rem 0; font-weight: bold;">0.68</p>
                        <p style="color: #BDDDFC; margin: 0; font-size: 0.9rem;">Global Average</p>
                    </div>
                    <div style="background: #2a3845; padding: 1rem; border-radius: 10px; text-align: center;">
                        <h4 style="color: #BDDDFC; margin: 0;">Top 10% Income Share</h4>
                        <p style="color: #BDDDFC; font-size: 1.5rem; margin: 0.5rem 0; font-weight: bold;">52%</p>
                        <p style="color: #BDDDFC; margin: 0; font-size: 0.9rem;">Of Total Income</p>
                    </div>
                    <div style="background: #2a3845; padding: 1rem; border-radius: 10px; text-align: center;">
                        <h4 style="color: #BDDDFC; margin: 0;">Countries Analyzed</h4>
                        <p style="color: #BDDDFC; font-size: 1.5rem; margin: 0.5rem 0; font-weight: bold;">156</p>
                        <p style="color: #BDDDFC; margin: 0; font-size: 0.9rem;">Worldwide</p>
                    </div>
                    <div style="background: #2a3845; padding: 1rem; border-radius: 10px; text-align: center;">
                        <h4 style="color: #BDDDFC; margin: 0;">Data Points</h4>
                        <p style="color: #BDDDFC; font-size: 1.5rem; margin: 0.5rem 0; font-weight: bold;">2.5M+</p>
                        <p style="color: #BDDDFC; margin: 0; font-size: 0.9rem;">Collected Since 1990</p>
                    </div>
                </div>
            </div>
        """, unsafe_allow_html=True)
    
    with col2:
        st.markdown("""
            <div style="background: #384959; 
                        padding: 2rem; border-radius: 15px; margin-bottom: 2rem;
                        border-left: 5px solid #BDDDFC; box-shadow: 0 5px 15px rgba(0, 0, 0, 0.2);">
                <h3 style="color: #BDDDFC; border-bottom: 2px solid #BDDDFC; padding-bottom: 0.5rem;">
                    🚀 Quick Actions
                </h3>
                <div style="display: flex; flex-direction: column; gap: 1rem;">
                    <div style="background: #2a3845; padding: 1rem; border-radius: 10px;">
                        <h4 style="color: #BDDDFC; margin: 0;">View Global Overview</h4>
                        <p style="color: #BDDDFC; margin: 0.5rem 0 0 0; font-size: 0.9rem;">Explore income inequality metrics worldwide</p>
                    </div>
                    <div style="background: #2a3845; padding: 1rem; border-radius: 10px;">
                        <h4 style="color: #BDDDFC; margin: 0;">Compare Regions</h4>
                        <p style="color: #BDDDFC; margin: 0.5rem 0 0 0; font-size: 0.9rem;">Analyze disparities between continents</p>
                    </div>
                    <div style="background: #2a3845; padding: 1rem; border-radius: 10px;">
                        <h4 style="color: #BDDDFC; margin: 0;">Download Reports</h4>
                        <p style="color: #BDDDFC; margin: 0.5rem 0 0 0; font-size: 0.9rem;">Export data for further analysis</p>
                    </div>
                </div>
            </div>
            
        """, unsafe_allow_html=True)

def show_profile():
    # Initialize session state for edit mode and form data
    if 'edit_mode' not in st.session_state:
        st.session_state.edit_mode = False
    if 'profile_data' not in st.session_state:
        # Initialize with default values
        st.session_state.profile_data = {
            'first_name': '',
            'last_name': '', 
            'email': st.session_state.get('email', ''),
            'contact_number': '',
            'profile_type': 'Student',
            'country': '',
            'profile_pic': None
        }
        # Then try to load existing data
        load_profile_data()

        st.markdown("""
        <style>
        .profile-container {
            display: flex;
            align-items: center;   /* Vertically center */
            justify-content: flex-start;
            gap: 2rem;             /* Space between pic and details */
            margin-top: 2rem;
            margin-bottom: 2rem;
        }
        .profile-pic {
            flex-shrink: 0;
        }
        .profile-details {
            flex: 1;
        }
        </style>
    """, unsafe_allow_html=True)
    
    # Debug button
    with st.sidebar:
        if st.button("🐛 Debug Profile"):
            st.write("### Profile Debug Info")
            st.write(f"Username: {st.session_state.username}")
            st.write(f"Profile Data: {st.session_state.profile_data}")
            st.write(f"Edit Mode: {st.session_state.edit_mode}")
            
            # Show what's in Excel
            try:
                profile_df = pd.read_excel(CREDENTIALS_FILE, sheet_name='profiles')
                st.write("### Excel Profiles Sheet")
                st.dataframe(profile_df)
            except Exception as e:
                st.error(f"Error reading Excel: {e}")

    # Header
    st.markdown("""
        <div style="background: #384959; padding: 2rem; border-radius: 15px; margin-bottom: 2rem; border: 2px solid #BDDDFC;">
            <h1 style="color: #BDDDFC; margin: 0; text-align: center;">👤 User Profile</h1>
        </div>
    """, unsafe_allow_html=True)

    # Edit/Save button at top
    col1, col2, col3 = st.columns([2, 1, 1])
    
    with col3:
        if not st.session_state.edit_mode:
            if st.button("✏️ Edit Profile", use_container_width=True):
                st.session_state.edit_mode = True
                st.rerun()
        else:
            col_save, col_cancel = st.columns(2)
            with col_save:
                if st.button("💾 Save Changes", use_container_width=True, type="primary"):
                    if save_profile_to_excel():
                        st.session_state.edit_mode = False
                        st.success("Profile updated successfully!")
                        st.rerun()
                    else:
                        st.error("Failed to save profile changes")
            with col_cancel:
                if st.button("❌ Cancel", use_container_width=True):
                    st.session_state.edit_mode = False
                    st.rerun()

    if st.session_state.edit_mode:
        show_edit_profile()
    else:
        show_view_profile()

def show_view_profile():
    """Display profile in view mode"""
    # Safe debugging - check if profile_data exists
    if 'profile_data' not in st.session_state:
        st.error("Profile data not initialized. Please log out and log in again.")
        return
    
    # Debug: Check what data we have
    print(f"Displaying profile for: {st.session_state.username}")
    print(f"Profile data: {st.session_state.profile_data}")
    
    col1, col2 = st.columns([1, 2])
    
    with col1:
        # Profile picture
        if st.session_state.profile_data.get('profile_pic'):
            st.image(st.session_state.profile_data['profile_pic'], width=150)
        else:
            # Default avatar with initial - FIXED: Handle empty first name
            first_name = st.session_state.profile_data.get('first_name', '')
            initial = first_name[0] if first_name else st.session_state.username[0] if st.session_state.username else 'U'
            st.markdown(f"""
                <div style="width: 150px; height: 150px; border-radius: 50%; background: #BDDDFC; 
                            display: flex; align-items: center; justify-content: center; margin: 0 auto;">
                    <span style="color: #384959; font-size: 3rem; font-weight: bold;">{initial.upper()}</span>
                </div>
            """, unsafe_allow_html=True)
        
        # Display profile type - FIXED: Use .get() with default
        #profile_type = st.session_state.profile_data.get('profile_type', 'Student')
        #st.markdown(f"<span style='color: #2C3E50; font-weight: bold; font-size: 16px;'>{profile_type}</span>", unsafe_allow_html=True)

    with col2:
        # Display profile information - FIXED: Handle empty names properly
        first_name = st.session_state.profile_data.get('first_name', '')
        last_name = st.session_state.profile_data.get('last_name', '')
        
        full_name = f"{first_name} {last_name}".strip()
        display_name = full_name if full_name else st.session_state.username if st.session_state.username else 'User'
        
        st.markdown(f"<h2 style='color: #2C3E50; border-bottom: 2px solid #BDDDFC; padding-bottom: 10px;'>{display_name}</h2>", unsafe_allow_html=True)
        
        st.markdown("---")
        
        info_col1, info_col2 = st.columns(2)
        
        with info_col1:
            st.markdown("<span style='color: #34495E; font-weight: bold; font-size: 16px;'>📧 Email:</span>", unsafe_allow_html=True)
            email = st.session_state.profile_data.get('email', st.session_state.get('email', 'Not provided'))
            st.markdown(f"<div style='color: #2C3E50; background-color: #ECF0F1; padding: 8px; border-radius: 5px; margin: 5px 0;'>{email}</div>", unsafe_allow_html=True)
            
            st.markdown("<span style='color: #34495E; font-weight: bold; font-size: 16px;'>📞 Contact:</span>", unsafe_allow_html=True)
            contact = st.session_state.profile_data.get('contact_number', '')
            contact_display = contact if contact else "Not provided"
            st.markdown(f"<div style='color: #2C3E50; background-color: #ECF0F1; padding: 8px; border-radius: 5px; margin: 5px 0;'>{contact_display}</div>", unsafe_allow_html=True)
        
        with info_col2:
            st.markdown("<span style='color: #34495E; font-weight: bold; font-size: 16px;'>🌍 Country:</span>", unsafe_allow_html=True)
            country = st.session_state.profile_data.get('country', '')
            country_display = country if country else "Not provided"
            st.markdown(f"<div style='color: #2C3E50; background-color: #ECF0F1; padding: 8px; border-radius: 5px; margin: 5px 0;'>{country_display}</div>", unsafe_allow_html=True)
            
            st.markdown("<span style='color: #34495E; font-weight: bold; font-size: 16px;'>👤 Type:</span>", unsafe_allow_html=True)
            profile_type = st.session_state.profile_data.get('profile_type', 'Student')
            st.markdown(f"<div style='color: #2C3E50; background-color: #ECF0F1; padding: 8px; border-radius: 5px; margin: 5px 0;'>{profile_type}</div>", unsafe_allow_html=True)

def show_edit_profile():
    """Display profile in edit mode"""
    col1, col2 = st.columns([1, 2])
    
    with col1:
        st.subheader("Profile Picture")
        
        # Profile picture upload
        uploaded_file = st.file_uploader("Upload Profile Picture", type=['jpg', 'jpeg', 'png'], key="profile_pic_upload")
        
        if uploaded_file is not None:
            st.session_state.profile_data['profile_pic'] = uploaded_file
            st.image(uploaded_file, width=150)
        elif st.session_state.profile_data['profile_pic']:
            st.image(st.session_state.profile_data['profile_pic'], width=150)
        else:
            initial = st.session_state.profile_data['first_name'][0] if st.session_state.profile_data['first_name'] else st.session_state.username[0]
            st.markdown(f"""
                <div style="width: 150px; height: 150px; border-radius: 50%; background: #BDDDFC; 
                            display: flex; align-items: center; justify-content: center; margin: 0 auto;">
                    <span style="color: #384959; font-size: 3rem; font-weight: bold;">{initial.upper()}</span>
                </div>
            """, unsafe_allow_html=True)
        
        # Profile type
        profile_types = ["Student", "Employee", "Researcher", "Professor", "Other"]
        current_index = profile_types.index(st.session_state.profile_data['profile_type']) if st.session_state.profile_data['profile_type'] in profile_types else 0
        st.session_state.profile_data['profile_type'] = st.selectbox(
            "Profile Type",
            profile_types,
            index=current_index
        )

    with col2:
        st.subheader("Personal Information")
        
        # Name fields
        col_name1, col_name2 = st.columns(2)
        with col_name1:
            st.session_state.profile_data['first_name'] = st.text_input(
                "First Name",
                value=st.session_state.profile_data['first_name'],
                placeholder="Enter your first name"
            )
        with col_name2:
            st.session_state.profile_data['last_name'] = st.text_input(
                "Last Name", 
                value=st.session_state.profile_data['last_name'],
                placeholder="Enter your last name"
            )
        
        # Email (read-only as it's from login)
        st.text_input(
            "Email Address",
            value=st.session_state.profile_data['email'],
            disabled=True,
            help="Email cannot be changed as it's your login identifier"
        )
        
        # Contact number
        st.session_state.profile_data['contact_number'] = st.text_input(
            "Contact Number",
            value=st.session_state.profile_data['contact_number'],
            placeholder="+1 234 567 8900"
        )
        
        # Country selection
        countries = [
            "", "United States", "United Kingdom", "Canada", "Australia", "Germany", "France", "Japan", 
            "China", "India", "Brazil", "South Africa", "Mexico", "Italy", "Spain", "South Korea",
            "Russia", "Netherlands", "Sweden", "Norway", "Denmark", "Finland", "Switzerland",
            "Austria", "Belgium", "Portugal", "Ireland", "New Zealand", "Singapore", "Malaysia",
            "Thailand", "Vietnam", "Philippines", "Indonesia", "Egypt", "Nigeria", "Kenya",
            "Saudi Arabia", "United Arab Emirates", "Qatar", "Israel", "Turkey", "Greece",
            "Poland", "Czech Republic", "Hungary", "Romania", "Ukraine", "Argentina", "Chile",
            "Colombia", "Peru", "Venezuela", "Other"
        ]
        
        current_country_index = countries.index(st.session_state.profile_data['country']) if st.session_state.profile_data['country'] in countries else 0
        st.session_state.profile_data['country'] = st.selectbox(
            "Country",
            countries,
            index=current_country_index
        )

def save_profile_to_excel():
    """Save profile data to Excel file"""
    try:
        # Load existing data
        try:
            profile_df = pd.read_excel(CREDENTIALS_FILE, sheet_name='profiles')
        except:
            profile_df = pd.DataFrame(columns=['username', 'first_name', 'last_name', 'email', 
                                             'contact_number', 'profile_type', 'country', 'last_updated'])
        
        # Update or add profile data
        if st.session_state.username in profile_df['username'].values:
            # Update existing
            profile_df.loc[profile_df['username'] == st.session_state.username, [
                'first_name', 'last_name', 'email', 'contact_number', 'profile_type', 'country', 'last_updated'
            ]] = [
                st.session_state.profile_data['first_name'],
                st.session_state.profile_data['last_name'],
                st.session_state.profile_data['email'],
                st.session_state.profile_data['contact_number'],
                st.session_state.profile_data['profile_type'],
                st.session_state.profile_data['country'],
                datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            ]
        else:
            # Add new
            new_profile = pd.DataFrame([{
                'username': st.session_state.username,
                'first_name': st.session_state.profile_data['first_name'],
                'last_name': st.session_state.profile_data['last_name'],
                'email': st.session_state.profile_data['email'],
                'contact_number': st.session_state.profile_data['contact_number'],
                'profile_type': st.session_state.profile_data['profile_type'],
                'country': st.session_state.profile_data['country'],
                'last_updated': datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            }])
            profile_df = pd.concat([profile_df, new_profile], ignore_index=True)
        
        # Load other sheets to preserve them
        user_df = load_user_credentials()
        feedback_df = load_feedback_data()
        
        # Save all sheets
        with pd.ExcelWriter(CREDENTIALS_FILE, engine='openpyxl') as writer:
            user_df.to_excel(writer, sheet_name='user_details', index=False)
            feedback_df.to_excel(writer, sheet_name='feedback', index=False)
            profile_df.to_excel(writer, sheet_name='profiles', index=False)
        
        return True
    except Exception as e:
        st.error(f"Error saving profile: {e}")
        return False

def load_profile_data():
    """Load profile data from Excel when user logs in"""
    try:
        # Try to load the profiles sheet
        profile_df = pd.read_excel(CREDENTIALS_FILE, sheet_name='profiles')
        
        # Check if the current user has a profile
        if st.session_state.username in profile_df['username'].values:
            user_profile = profile_df[profile_df['username'] == st.session_state.username].iloc[0]
            
            # Update profile data with values from Excel
            st.session_state.profile_data = {
                'first_name': user_profile['first_name'] if pd.notna(user_profile['first_name']) else '',
                'last_name': user_profile['last_name'] if pd.notna(user_profile['last_name']) else '',
                'email': st.session_state.email,  # Use current email from login
                'contact_number': user_profile['contact_number'] if pd.notna(user_profile['contact_number']) else '',
                'profile_type': user_profile['profile_type'] if pd.notna(user_profile['profile_type']) else 'Student',
                'country': user_profile['country'] if pd.notna(user_profile['country']) else '',
                'profile_pic': None
            }
        else:
            # If no profile exists, initialize with default values
            st.session_state.profile_data = {
                'first_name': '',
                'last_name': '', 
                'email': st.session_state.email,
                'contact_number': '',
                'profile_type': 'Student',
                'country': '',
                'profile_pic': None
            }
    except Exception as e:
        # If any error occurs (file doesn't exist, sheet doesn't exist, etc.)
        # Initialize with default values
        st.session_state.profile_data = {
            'first_name': '',
            'last_name': '', 
            'email': st.session_state.email,
            'contact_number': '',
            'profile_type': 'Student',
            'country': '',
            'profile_pic': None
        }   

def show_about():
    st.markdown("""
        <div style="background: #384959; 
                    padding: 2.5rem; border-radius: 20px; margin-bottom: 2rem; 
                    box-shadow: 0 8px 25px rgba(0, 0, 0, 0.4);">
            <h1 style="color: #BDDDFC; margin: 0; text-align: center;">📊 About the Dashboard</h1>
            <p style="color: #BDDDFC; font-size: 1.2rem; margin: 0.5rem 0 0 0; text-align: center;">
                Learn about global income inequality visualization
            </p>
        </div>
    """, unsafe_allow_html=True)
    
    # Shared CSS for equal box sizing
    st.markdown("""
        <style>
        .equal-box {
            height: 100%;
            display: flex;
            flex-direction: column;
            justify-content: center;
        }
        .about-image {
            width: 100% !important;
            height: 100% !important;
            object-fit: cover;
            border-radius: 15px;
            box-shadow: 0 5px 15px rgba(0, 0, 0, 0.2);
            margin-bottom: 2rem;
        }
        </style>
    """, unsafe_allow_html=True)

    # Section 1: Project Overview
    col1, col2 = st.columns([2, 1])
    with col1:
        st.markdown("""
            <div class="equal-box" style="background: #384959; 
                        padding: 2rem; border-radius: 15px; margin-bottom: 2rem;
                        border-left: 5px solid #BDDDFC; box-shadow: 0 5px 15px rgba(0, 0, 0, 0.2);">
                <h2 style="color: #BDDDFC; border-bottom: 2px solid #BDDDFC; padding-bottom: 0.5rem;">
                    Project Overview
                </h2>
                <p style="color: #BDDDFC; line-height: 1.7;">
                    This interactive dashboard visualizes global income inequality data from various reliable sources, 
                    allowing researchers, policymakers, and the general public to explore trends and patterns in income 
                    distribution across countries and regions.
                </p>
            </div>
        """, unsafe_allow_html=True)
    with col2:
        st.markdown('<div class="equal-box">', unsafe_allow_html=True)
        st.image("https://bigthink.com/wp-content/uploads/2022/12/AdobeStock_503474263.jpeg", use_container_width=True)
        st.markdown('</div>', unsafe_allow_html=True)

    # Section 2: Key Features
    col1, col2 = st.columns([2, 1])
    with col1:
        st.markdown("""
            <div class="equal-box" style="background: #384959; 
                        padding: 2rem; border-radius: 15px; margin-bottom: 2rem;
                        border-left: 5px solid #BDDDFC; box-shadow: 0 5px 15px rgba(0, 0, 0, 0.2);">
                <h3 style="color: #BDDDFC; border-bottom: 2px solid #BDDDFC; padding-bottom: 0.5rem;">
                    🎯 Key Features
                </h3>
                <ul style="color: #BDDDFC; line-height: 1.7;">
                    <li><strong style="color: #BDDDFC;">Global Inequality Overview</strong>: Explore income inequality metrics across different countries</li>
                    <li><strong style="color: #BDDDFC;">Income Distribution Analysis</strong>: Dive deep into income distribution patterns and Gini coefficients</li>
                    <li><strong style="color: #BDDDFC;">Regional Comparison</strong>: Compare income inequality trends across different regions</li>
                </ul>
            </div>
        """, unsafe_allow_html=True)
    with col2:
        st.markdown('<div class="equal-box">', unsafe_allow_html=True)
        st.image("https://www.shutterstock.com/image-illustration/income-inequality-wealth-distribution-art-260nw-135704756.jpg", use_container_width=True)
        st.markdown('</div>', unsafe_allow_html=True)

    # Section 3: Data Sources
    col1, col2 = st.columns([2, 1])
    with col1:
        st.markdown("""
            <div class="equal-box" style="background: #384959; 
                        padding: 2rem; border-radius: 15px; margin-bottom: 2rem;
                        border-left: 5px solid #BDDDFC; box-shadow: 0 5px 15px rgba(0, 0, 0, 0.2);">
                <h3 style="color: #BDDDFC; border-bottom: 2px solid #BDDDFC; padding-bottom: 0.5rem;">
                    📚 Data Sources
                </h3>
                <p style="color: #BDDDFC; line-height: 1.7;">
                    Our data is compiled from reputable sources including:
                </p>
                <ul style="color: #BDDDFC; line-height: 1.7;">
                    <li>World Bank Development Indicators</li>
                    <li>OECD Income Distribution Database</li>
                    <li>World Inequality Database</li>
                    <li>UN Human Development Reports</li>
                </ul>
            </div>
        """, unsafe_allow_html=True)
    with col2:
        st.markdown('<div class="equal-box">', unsafe_allow_html=True)
        st.image("https://acquiscent.com/wp-content/uploads/2021/01/rsz_depositphotos_64676541_s-2019.jpg", use_container_width=True)
        st.markdown('</div>', unsafe_allow_html=True)


def show_dashboard():
    # Custom Styling with the custom blue color palette
    st.markdown("""
        <style>
            /* Main background with beautiful gradient - KEPT INTACT */
            .stApp {
                background: linear-gradient(135deg, #6A89A7 0%, #88BDF2 50%, #BDDDFC 100%);
                background-attachment: fixed;
                background-size: cover;
            }
            
            /* Main content area - CHANGED BOX COLOR */
            .main .block-container {
                background: #384959;
                border-radius: 20px;
                padding: 3rem;
                margin-top: 1.5rem;
                margin-bottom: 1.5rem;
                box-shadow: 0 20px 40px rgba(0, 0, 0, 0.15);
                border: 2px solid #BDDDFC;
            }
            
            /* Sidebar styling - CHANGED BOX COLOR AND SIZE */
            section[data-testid="stSidebar"] {
                background: #384959 !important;
                min-width: 200px !important;
                max-width: 280px !important;
            }
            
            /* Reduce sidebar width */
            [data-testid="stSidebar"] > div:first-child {
                width: 250px !important;
            }
            
            /* Make sidebar content more compact */
            .sidebar-content {
                padding: 1rem 0.5rem !important;
            }
            
            /* Reduce font sizes in sidebar */
            section[data-testid="stSidebar"] * {
                color: #BDDDFC !important;
                font-weight: 500;
                font-size: 14px !important;
            }
            
            /* Smaller sidebar headers */
            .sidebar-header {
                color: #BDDDFC !important;
                font-weight: 700 !important;
                font-size: 16px !important;
            }
            
            /* Smaller buttons in sidebar */
            .stButton > button {
                background: #384959;
                color: #BDDDFC;
                border: 2px solid #BDDDFC;
                border-radius: 8px;
                padding: 0.5rem 1rem !important;
                font-weight: 600;
                transition: all 0.3s ease;
                box-shadow: 0 3px 10px rgba(0, 0, 0, 0.3);
                font-size: 14px !important;
                margin: 0.2rem 0 !important;
            }
            
            .stButton > button:hover {
                background: #2a3845;
                transform: translateY(-1px);
                box-shadow: 0 5px 15px rgba(0, 0, 0, 0.4);
                color: #BDDDFC;
                border: 2px solid #BDDDFC;
            }
            
            /* Reduce padding in sidebar sections */
            .stSidebar div[data-testid="stVerticalBlock"] {
                gap: 0.3rem !important;
            }
            
            /* Smaller user info section */
            .user-info {
                padding: 0.5rem !important;
                margin-bottom: 0.5rem !important;
            }
            
            /* Compact navigation items */
            .nav-item {
                margin: 0.1rem 0 !important;
            }
            
            /* Reduce spacing around dividers */
            hr {
                margin: 0.8rem 0 !important;
                border-color: #BDDDFC !important;
            }
            
            /* Smaller select boxes in sidebar */
            .stSelectbox > div > div {
                background: #384959;
                border: 1px solid #BDDDFC;
                border-radius: 6px;
                color: #BDDDFC;
                font-size: 14px !important;
                padding: 0.3rem !important;
            }
            
            /* Rest of your existing styles... */
            .stTextInput > div > div > input {
                border-radius: 12px;
                border: 2px solid #BDDDFC;
                padding: 0.9rem;
                transition: all 0.3s ease;
                background: #384959;
                font-size: 1rem;
                color: #BDDDFC;
            }
            
            .dashboard-header {
                background: #384959;
                color: #BDDDFC;
                padding: 2rem;
                border-radius: 20px;
                text-align: center;
                margin-bottom: 2rem;
                box-shadow: 0 15px 35px rgba(0, 0, 0, 0.3);
                border: 2px solid #BDDDFC;
            }
            
            /* ... keep the rest of your styles ... */
        </style>
    """, unsafe_allow_html=True)

    # Header Section
    st.markdown("""
        <div class="dashboard-header">
            <h1 style="margin: 0; font-size: 1.1rem; text-shadow: 0 2px 4px rgba(0,0,0,0.1);">
                VISIONARY - INTERACTIVE DASHBOARDS
            </h1>
            <p style="font-size: 1.3rem; margin: 1rem 0 0 0; font-weight: 400;">
                Explore interactive Power BI dashboards with beautiful data insights
            </p>
        </div>
    """, unsafe_allow_html=True)

    # User info in sidebar
    st.sidebar.markdown("---")
    st.sidebar.markdown(f"### 👤 Welcome, {st.session_state.username}!")
    st.sidebar.markdown(f"📧 {st.session_state.email}")
    
    # Navigation in sidebar - ADDED NEW BUTTONS
    st.sidebar.markdown("---")
    st.sidebar.markdown("### 🧭 Navigation")
    
    # Navigation options - ADDED HOME, PROFILE, FEEDBACK
    nav_options = {
        "🏠 Home": "home",
        "📊 Dashboard": "dashboard",
        "📋 About": "about",
        "🤖 ChatBot": "chatbot",
        "👤 Profile": "profile",
        "💬 Feedback": "feedback"
        }
    
    for option, key in nav_options.items():
        if st.sidebar.button(option, key=key, use_container_width=True):
            st.session_state.current_nav = key
    
    if st.sidebar.button("🚪 Logout", use_container_width=True):
        st.session_state.logged_in = False
        st.rerun()

    # Sidebar Dashboard Selector
    st.sidebar.markdown("---")
    st.sidebar.header("📈 Select a Dashboard")

    dashboard_links = {
        "🌐 Global Inequality Overview": "https://app.powerbi.com/view?r=eyJrIjoiZjIxZmMyNjktMzRiZC00NmNlLWFhZjAtZmI0ODk0ODBhYWEzIiwidCI6IjAzMzNmNGQwLTNhZGYtNGFjMC05YTYxLTJmYjNjM2NkMDk0NiJ9",
    }

    selected_dashboard = st.sidebar.selectbox(
        "Choose Dashboard",
        list(dashboard_links.keys()),
        label_visibility="collapsed"
    )

    # Display content based on navigation - ADDED NEW PAGES
    if st.session_state.get('current_nav') == 'home':
        show_home()
    elif st.session_state.get('current_nav') == 'profile':
        show_profile()
    elif st.session_state.get('current_nav') == 'feedback':
        show_feedback()
    elif st.session_state.get('current_nav') == 'about':
        show_about()
    elif st.session_state.get('current_nav') == 'chatbot':
        show_chatbot()
    else:
        # Dashboard description with beautiful styling
        dashboard_descriptions = {
            "🌐 Global Inequality Overview": "Explore income inequality metrics across different countries and regions with interactive maps and charts.",
        }
        
        st.markdown(f"""
            <div class="info-card">
                <h2 style="margin-bottom: 1rem; border-bottom: 3px solid #BDDDFC; padding-bottom: 0.5rem;">
                    {selected_dashboard}
                </h2>
                <p class="beautiful-text" style="font-size: 1.1rem;">
                    {dashboard_descriptions[selected_dashboard]}
                </p>
            </div>
        """, unsafe_allow_html=True)
        
        st.markdown("")

        # Display Power BI Dashboard
        st.markdown(f"""
            <div class="iframe-container">
                <iframe 
                    src="{dashboard_links[selected_dashboard]}" 
                    width="100%" 
                    height="850" 
                    frameborder="0" 
                    allowFullScreen="true">
                </iframe>
            </div>
        """, unsafe_allow_html=True)

def main():
    # Initialize session state
    initialize_excel()
    if 'logged_in' not in st.session_state:
        st.session_state.logged_in = False
    if 'username' not in st.session_state:
        st.session_state.username = None
    if 'email' not in st.session_state:
        st.session_state.email = None
    if 'current_nav' not in st.session_state:
        st.session_state.current_nav = 'home'
    if 'edit_mode' not in st.session_state:
        st.session_state.edit_mode = False
    if 'profile_data' not in st.session_state:
        st.session_state.profile_data = {
            'first_name': '',
            'last_name': '', 
            'email': '',
            'contact_number': '',
            'profile_type': 'Student',
            'country': '',
            'profile_pic': None
        }

    if st.session_state.logged_in:
        # Ensure profile data is loaded when user is logged in
        if st.session_state.username and not st.session_state.profile_data.get('email'):
            load_profile_data()
        show_dashboard()
    else:
        # Auth section with beautiful styling - ADDED BACKGROUND IMAGE
        st.markdown("""
            <style>
                .stApp {
                    background: 
                        url("https://confidential-white-885jel0oj5.edgeone.app/world_night.jpg");
                    background-size: cover;
                    background-position: center;
                }
                .auth-container {
                    background: rgba(56, 73, 89, 0.9);
                    padding: 3rem;
                    border-radius: 20px;
                    margin: 2rem auto;
                    box-shadow: 0 20px 40px rgba(0, 0, 0, 0.4);
                    border: 2px solid #BDDDFC;
                }
                .auth-header {
                    text-align: center;
                    margin-bottom: 2rem;
                }
                .auth-header h1 {
                    color: #BDDDFC;
                    font-size: 5.0rem;
                    margin-bottom: 0.5rem;
                    font-weight: 800;
                }
                .auth-header p {
                    color: #BDDDFC;
                    font-size: 1.2rem;
                }
            </style>
        """, unsafe_allow_html=True)
        
        # Center the auth form with background
        col1, col2, col3 = st.columns([1, 2, 1])
        
        with col2:
            st.markdown("""
                <div class="auth-container">
                    <div class="auth-header">
                        <h1>🌍 Global Income Inequality Dashboard</h1>
                        <p>Visualizing Global Income Inequality with Beautiful Insights</p>
                    </div>
            """, unsafe_allow_html=True)
            
            # Auth option tabs
            tab1, tab2 = st.tabs(["🔑 Login", "🔐 Sign Up"])
            
            with tab1:
                show_login()
                
            with tab2:
                show_signup()
                
            st.markdown("</div>", unsafe_allow_html=True)

if __name__ == '__main__':
    main()