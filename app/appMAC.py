import streamlit as st
import pandas as pd
from canvasapi import Canvas
from datetime import datetime, timezone
import io
import re
from openpyxl.styles import Font, PatternFill, Alignment

# --- PAGE CONFIG ---
st.set_page_config(page_title="Canvas Planner", page_icon="📅", layout="wide")

st.markdown("""
    <style>
    .main { background-color: #f5f5f7; }
    .stButton>button { width: 100%; border-radius: 10px; height: 3em; background-color: #0071e3; color: white; }
    </style>
    """, unsafe_allow_globals=True)

st.title("📅 Canvas Assignment Planner")
st.write("Enter your details in the sidebar to generate your local Excel schedule.")

# --- SIDEBAR ---
st.sidebar.header("Canvas Credentials")
canvas_url = st.sidebar.text_input("Canvas URL", value="https://vsu.instructure.com")
api_key = st.sidebar.text_input("API Key", type="password", help="Get this from Canvas > Account > Settings > New Access Token")

# --- APP LOGIC ---
if st.sidebar.button("🚀 Fetch Assignments"):
    if not api_key:
        st.error("Please enter your API Key!")
    else:
        with st.spinner("Connecting to Canvas and gathering courses..."):
            try:
                canvas = Canvas(canvas_url, api_key)
                user = canvas.get_current_user()
                courses = user.get_courses(enrollment_state='active')
                
                all_data = []
                now = datetime.now(timezone.utc)

                for course in courses:
                    try:
                        c_name = getattr(course, 'name', 'Unknown Course')
                        if any(x in c_name for x in ["Orientation", "Testing", "Support"]): continue
                        
                        assignments = course.get_assignments()
                        for a in assignments:
                            due_at = getattr(a, 'due_at', None)
                            
                            # Calculate Days Left
                            days_left = "N/A"
                            if due_at:
                                due_date_obj = pd.to_datetime(due_at).tz_localize(None).tz_localize(timezone.utc)
                                days_left = (due_date_obj - now).days

                            all_data.append({
                                "Status": "✅ Done" if getattr(a, "has_submitted_submissions", False) else "⏳ Pending",
                                "Course": c_name,
                                "Assignment": getattr(a, "name", "Unnamed"),
                                "Days Left": days_left,
                                "Due Date": pd.to_datetime(due_at).strftime('%b %d, %I:%M %p') if due_at else "—",
                                "Link": getattr(a, "html_url", "")
                            })
                    except: continue

                if all_data:
                    df = pd.DataFrame(all_data)
                    st.success(f"Found {len(df)} assignments!")
                    st.dataframe(df, use_container_width=True)

                    # --- EXCEL EXPORT WITH STYLING ---
                    output = io.BytesIO()
                    with pd.ExcelWriter(output, engine='openpyxl') as writer:
                        df.to_excel(writer, index=False, sheet_name='My Schedule')
                        workbook = writer.book
                        worksheet = writer.sheets['My Schedule']
                        
                        # Style Headers
                        header_fill = PatternFill(start_color="334155", end_color="334155", fill_type="solid")
                        header_font = Font(color="FFFFFF", bold=True)
                        for cell in worksheet[1]:
                            cell.fill = header_fill
                            cell.font = header_font
                            cell.alignment = Alignment(horizontal='center')
                        
                        # Auto-adjust column width
                        for col in worksheet.columns:
                            max_length = 0
                            column = col[0].column_letter
                            for cell in col:
                                try:
                                    if len(str(cell.value)) > max_length: max_length = len(str(cell.value))
                                except: pass
                            worksheet.column_dimensions[column].width = max_length + 2

                    st.download_button(
                        label="📥 Download Styled Excel Sheet",
                        data=output.getvalue(),
                        file_name="Canvas_Planner.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                else:
                    st.warning("No assignments found in your active courses.")
            except Exception as e:
                st.error(f"Connection Failed: {e}")