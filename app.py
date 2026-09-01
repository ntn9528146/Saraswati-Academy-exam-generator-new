import streamlit as st
import json
import os
import importlib
from dotenv import load_dotenv

load_dotenv()

st.set_page_config(
    page_title="School Faculty Cloud Portal",
    page_icon="🎓",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ----------------- MODERN UI STYLING -----------------
st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@400;500;600;700;800&display=swap');
    html, body, [class*="css"] {
        font-family: 'Plus+Jakarta+Sans', sans-serif;
    }
</style>
""", unsafe_allow_html=True)

# ----------------- USER DATABASE -----------------
DB_FILE = "teacher_users.json"
DEFAULT_USERS = {
    "TCH101": {
        "password": "password123",
        "name": "Nitin Tripathi",
        "dob": "1994-08-15",
        "gender": "Male",
        "subject": "Computer Science & IT (083/402)",
        "phone": "+91 9876543210",
        "email": "nitin.tripathi@school.edu.in",
        "qualification": "MCA / B.Ed",
        "photo": None
    },
    "TCH102": {
        "password": "password123",
        "name": "Pooja Sharma",
        "dob": "1992-04-10",
        "gender": "Female",
        "subject": "Mathematics (041)",
        "phone": "+91 9876501234",
        "email": "pooja.sharma@school.edu.in",
        "qualification": "M.Sc Mathematics",
        "photo": None
    }
}

def load_users():
    if os.path.exists(DB_FILE):
        try:
            with open(DB_FILE, "r") as f:
                return json.load(f)
        except Exception:
            return DEFAULT_USERS
    return DEFAULT_USERS

def save_users(data):
    with open(DB_FILE, "w") as f:
        json.dump(data, f, indent=4)

if "users_db" not in st.session_state:
    st.session_state["users_db"] = load_users()

if "logged_in_user" not in st.session_state:
    st.session_state["logged_in_user"] = None

# Default Logo Check
DEFAULT_LOGO_FILES = ["school_logo.png", "SA-Logo.png", "logo.png", "school_logo.jpg"]
default_logo_path = None
for fname in DEFAULT_LOGO_FILES:
    if os.path.exists(fname):
        default_logo_path = fname
        break

# Backend API Key
api_key = st.secrets.get("GEMINI_API_KEY", os.getenv("GEMINI_API_KEY", None))

# =====================================================================
#                          LOGIN PAGE
# =====================================================================
if st.session_state["logged_in_user"] is None:
    col_l1, col_l2, col_l3 = st.columns([1, 1.8, 1])
    with col_l2:
        st.markdown("<br>", unsafe_allow_html=True)
        if default_logo_path and os.path.exists(default_logo_path):
            st.image(default_logo_path, width=110)
        
        st.markdown("## 🔐 Faculty & Teacher Portal")
        st.caption("Authorized Examination & Academic Suite")
        
        with st.form("login_form"):
            t_id = st.text_input("Teacher ID / Username", placeholder="e.g. TCH101").strip().upper()
            t_pwd = st.text_input("Password", type="password", placeholder="Enter password")
            btn_login = st.form_submit_button("🚀 Sign In to Portal", use_container_width=True, type="primary")

            if btn_login:
                users = st.session_state["users_db"]
                if t_id in users and users[t_id]["password"] == t_pwd:
                    st.session_state["logged_in_user"] = t_id
                    st.success("✅ Login successful!")
                    st.rerun()
                else:
                    st.error("❌ Invalid ID or Password.")
    st.stop()

# =====================================================================
#             LOGGED-IN DASHBOARD & APP NAVIGATION
# =====================================================================
current_user_id = st.session_state["logged_in_user"]
user_data = st.session_state["users_db"][current_user_id]

# Top Modern Navbar
col_nav1, col_nav2 = st.columns([3, 1])
with col_nav1:
    st.markdown(
        f"""
        <div style="background: linear-gradient(135deg, #1e293b, #0f172a); padding: 16px 22px; border-radius: 12px; color: white; margin-bottom: 20px;">
            <h3 style="margin:0; font-size: 20px; color: #38bdf8;">🏫 SARASWATI ACADEMY</h3>
            <p style="margin:4px 0 0 0; font-size: 14px; color: #cbd5e1;">
                👋 Welcome, <b>{user_data['name']}</b> &nbsp;|&nbsp; 
                <span style="background: #0284c7; padding: 2px 8px; border-radius: 6px; font-size: 12px;">ID: {current_user_id}</span> &nbsp;|&nbsp; 
                <span>Dept: {user_data.get('subject', 'Faculty')}</span>
            </p>
        </div>
        """,
        unsafe_allow_html=True
    )
with col_nav2:
    st.write("")
    if st.button("🚪 Logout", type="secondary", use_container_width=True):
        st.session_state["logged_in_user"] = None
        st.rerun()

# ----------------- SIDEBAR: LOGO & MODULE SELECTOR -----------------
with st.sidebar:
    st.header("🏫 Academic Tools")
    if default_logo_path:
        st.image(default_logo_path, caption="Active School Logo", width=130)
    
    uploaded_logo = st.file_uploader("Override Logo (Optional)", type=["png", "jpg", "jpeg"])
    logo_temp_path = default_logo_path
    if uploaded_logo:
        with open("temp_logo.png", "wb") as f:
            f.write(uploaded_logo.getbuffer())
        logo_temp_path = "temp_logo.png"

    st.markdown("---")
    # Dynamic list of modules
    active_nav = st.radio(
        "📌 Choose Work Area:",
        [
            "📝 Exam Paper Generator",
            "👤 My Teacher Profile",
            "⚙️ Account & Security"
        ]
    )

# ----------------- MODULE ROUTER -----------------
if active_nav == "👤 My Teacher Profile":
    st.subheader("👤 Faculty Profile Details")
    col_p1, col_p2 = st.columns([1, 2.5])
    with col_p1:
        if user_data.get("photo") and os.path.exists(user_data["photo"]):
            st.image(user_data["photo"], width=170, caption=user_data['name'])
        else:
            st.info("📸 No profile photo set.")
        
        new_photo = st.file_uploader("Upload Profile Photo", type=["jpg", "png", "jpeg"], key="prof_pic")
        if new_photo:
            os.makedirs("user_photos", exist_ok=True)
            p_path = f"user_photos/{current_user_id}.png"
            with open(p_path, "wb") as f:
                f.write(new_photo.getbuffer())
            st.session_state["users_db"][current_user_id]["photo"] = p_path
            save_users(st.session_state["users_db"])
            st.success("✅ Profile photo updated!")
            st.rerun()

    with col_p2:
        with st.form("profile_edit_form"):
            col_d1, col_d2 = st.columns(2)
            with col_d1:
                name_val = st.text_input("Full Name", value=user_data.get("name", ""))
                dob_val = st.text_input("Date of Birth", value=user_data.get("dob", "1995-01-01"))
                gender_val = st.selectbox("Gender", ["Male", "Female", "Other"], index=0 if user_data.get("gender") == "Male" else 1)
            with col_d2:
                sub_val = st.text_input("Assigned Subject", value=user_data.get("subject", ""))
                phone_val = st.text_input("Contact Phone", value=user_data.get("phone", ""))
                email_val = st.text_input("Official Email", value=user_data.get("email", ""))
            
            qual_val = st.text_input("Educational Qualification", value=user_data.get("qualification", ""))
            if st.form_submit_button("💾 Save Profile Changes", type="primary"):
                st.session_state["users_db"][current_user_id].update({
                    "name": name_val, "dob": dob_val, "gender": gender_val,
                    "subject": sub_val, "phone": phone_val, "email": email_val,
                    "qualification": qual_val
                })
                save_users(st.session_state["users_db"])
                st.success("🎉 Profile updated!")
                st.rerun()

elif active_nav == "⚙️ Account & Security":
    st.subheader("🔒 Change Account Password")
    with st.form("pwd_change_form"):
        old_p = st.text_input("Current Password", type="password")
        new_p = st.text_input("New Password", type="password")
        confirm_p = st.text_input("Confirm New Password", type="password")
        if st.form_submit_button("🔑 Update Password", type="primary"):
            if user_data["password"] != old_p:
                st.error("❌ Current password is incorrect.")
            elif len(new_p) < 6:
                st.error("❌ Minimum 6 characters required.")
            elif new_p != confirm_p:
                st.error("❌ New passwords do not match.")
            else:
                st.session_state["users_db"][current_user_id]["password"] = new_p
                save_users(st.session_state["users_db"])
                st.success("✅ Password successfully updated!")

elif active_nav == "📝 Exam Paper Generator":
    # Exam Generator Module ko call karein
    from modules.exam_generator import render_exam_generator
    render_exam_generator(api_key=api_key, default_logo_path=logo_temp_path, default_school_name="Saraswati Academy")
