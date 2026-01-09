import streamlit as st
import pandas as pd
from canvasapi import Canvas
import io

st.set_page_config(page_title="Canvas to Excel", page_icon="📊")

st.title("🎓 Canvas Assignment Downloader")
st.write("Enter your details below to export your assignments to Excel.")

# User Inputs
base_url = st.text_input("Canvas URL", value="https://canvas.instructure.com")
api_token = st.text_input("API Token", type="password", help="Get this from your Canvas Settings > Approved Integrations")

if st.button("Generate Excel File"):
    if not api_token:
        st.error("Please enter an API Token!")
    else:
        with st.spinner("Fetching data from Canvas..."):
            try:
                # 1. Connect to Canvas
                canvas = Canvas(base_url, api_token)
                user = canvas.get_current_user()
                
                # 2. Your Logic (Example)
                courses = user.get_courses(enrollment_state='active')
                data = []
                for course in courses:
                    if hasattr(course, 'name'):
                        data.append({"Course": course.name, "ID": course.id})
                
                df = pd.DataFrame(data)

                # 3. Create Excel file in memory
                buffer = io.BytesIO()
                with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                    df.to_excel(writer, index=False)
                
                # 4. Show Download Button
                st.success(f"Hello {user.name}! Data ready.")
                st.download_button(
                    label="📥 Download Excel",
                    data=buffer.getvalue(),
                    file_name="canvas_assignments.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            except Exception as e:
                st.error(f"Connection failed: {e}")