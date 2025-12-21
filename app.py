from dotenv import load_dotenv

load_dotenv()

import streamlit as st
import pandas as pd
import json
import base64
from groq import Groq
from openpyxl import load_workbook
from PIL import Image
import io

# ================= CONFIG ================= #
MAX_IMAGES = 20

TARGET_FIELDS = [
    "Workshop Code",
    "Workshop Name",
    "Location",
    "Claim No",
    "Dealer Lot",
    "Mileage",
    "Repair Date",
    "Registration Date",
    "Casual Part No",
    "Casual Part Name",
    "VIN Last 6",
    "Model Code",
    "Dealer Observation"
]

# ================= INIT ================= #
client = Groq(api_key=st.secrets["GROQ_API_KEY"])

# ================= LOCAL TESTING ================= #
# client = Groq()

# ================= PROMPT ================= #
def build_prompt():
    return f"""
You are a vision-based information extraction engine.

Extract information from the image and return ONLY valid JSON.
Do NOT include explanations.
Do NOT include markdown.
Do NOT hallucinate values.
If a field is missing or unclear, return an empty string.

TARGET SCHEMA (keys must match EXACTLY):

{json.dumps({k: "" for k in TARGET_FIELDS}, indent=2)}

RULES:
- extract direct place/city name for the location (e.g., "PUNALUR-SRV, PUNALUR" -> "PUNALUR")
- Dates must be dd-mm-yy
- Mileage must be an integer
- VIN Last 6 must be numeric
- Dealer Observation must be UPPERCASE
"""

# ================= IMAGE → STRUCTURED ================= #
def image_to_structured(image_bytes):
    image_b64 = base64.b64encode(image_bytes).decode()

    completion = client.chat.completions.create(
        model="meta-llama/llama-4-maverick-17b-128e-instruct",
        messages=[
            {
                "role": "user",
                "content": [
                    {"type": "text", "text": build_prompt()},
                    {
                        "type": "image_url",
                        "image_url": {
                            "url": f"data:image/png;base64,{image_b64}"
                        }
                    }
                ]
            }
        ],
        temperature=0,
        max_completion_tokens=1024,
        stream=True
    )

    full_response = ""
    for chunk in completion:
        delta = chunk.choices[0].delta
        if delta and delta.content:
            full_response += delta.content

    return json.loads(full_response)

# ================= STREAMLIT UI ================= #
st.set_page_config(
    page_title="Vision Form Extraction (Groq Maverick)",
    layout="wide"
)

st.title("📸 Image → Structured Data → Excel (Groq Maverick)")

images = st.file_uploader(
    "Upload form images (max 20)",
    type=["png", "jpg", "jpeg"],
    accept_multiple_files=True
)

excel = st.file_uploader(
    "Upload Excel file (.xlsx)",
    type=["xlsx"]
)

if images and len(images) > MAX_IMAGES:
    st.error("Maximum 20 images allowed")
    st.stop()

if not images or not excel:
    st.info("Upload images and an Excel file to continue")
    st.stop()

if "uploaded_excel_name" not in st.session_state:
    st.session_state["uploaded_excel_name"] = None
    st.session_state["excel_bytes_cache"] = None
    st.session_state["updated_workbook"] = None
    st.session_state["updated_sheet"] = None
    st.session_state["update_message"] = None

if st.session_state["uploaded_excel_name"] != excel.name:
    st.session_state["uploaded_excel_name"] = excel.name
    st.session_state["excel_bytes_cache"] = excel.getvalue()
    st.session_state["updated_workbook"] = None
    st.session_state["updated_sheet"] = None
    st.session_state["update_message"] = None
elif st.session_state["excel_bytes_cache"] is None:
    st.session_state["excel_bytes_cache"] = excel.getvalue()

excel_bytes = st.session_state["excel_bytes_cache"]

preview_buffer = io.BytesIO(excel_bytes)
wb_preview = load_workbook(preview_buffer, read_only=True)
excel_sheetnames = wb_preview.sheetnames
wb_preview.close()
preview_buffer.close()

if not excel_sheetnames:
    st.error("No worksheets found in the uploaded workbook")
    st.stop()

records = []
image_map = {}
failed = []

# ================= PROCESS ================= #
with st.spinner("Processing images with Groq Maverick Vision Model..."):
    for idx, img in enumerate(images):
        try:
            img_bytes = img.read()
            structured = image_to_structured(img_bytes)

            record = structured.copy()
            record["Row ID"] = idx
            record["Source Image"] = img.name

            records.append(record)
            image_map[idx] = img_bytes

        except Exception:
            failed.append(img.name)

# ================= SUMMARY ================= #
st.subheader("📊 Processing Summary")
st.success(f"Extracted records: {len(records)}")
st.error(f"Failed images: {len(failed)}")

if failed:
    with st.expander("❌ Failed Images"):
        for f in failed:
            st.write(f)

# ================= REVIEW TABLE ================= #
if records:
    df = pd.DataFrame(records)
    ordered_cols = ["Row ID", "Source Image"] + TARGET_FIELDS
    df = df[ordered_cols]

    st.subheader("✏️ Review & Edit Extracted Data")

    edited_df = st.data_editor(
        df,
        num_rows="dynamic",
        use_container_width=True
    )

    # ================= IMAGE PREVIEW ================= #
    st.subheader("🖼️ Preview Source Image")

    selected_row = st.selectbox(
        "Select Row ID to preview image",
        options=edited_df["Row ID"].tolist()
    )

    if selected_row in image_map:
        st.image(
            image_map[selected_row],
            caption=f"Source Image for Row ID {selected_row}",
            use_column_width=True
        )

    st.subheader("📎 Excel Destination")
    selected_sheet = st.selectbox(
        "Select worksheet to append rows",
        options=excel_sheetnames,
        help="All reviewed rows will be appended to this sheet."
    )

    # ================= SUBMIT ================= #
    if st.button("✅ Submit to Excel"):
        with st.spinner("Updating workbook..."):
            workbook_buffer = io.BytesIO(excel_bytes)
            wb = load_workbook(workbook_buffer)

            if selected_sheet not in wb.sheetnames:
                st.error("Selected worksheet is not available in the uploaded workbook")
            else:
                ws = wb[selected_sheet]

                for _, row in edited_df.iterrows():
                    ws.append(
                        [row[col] for col in TARGET_FIELDS]
                    )

                output_buffer = io.BytesIO()
                wb.save(output_buffer)
                output_buffer.seek(0)
                updated_bytes = output_buffer.getvalue()
                output_buffer.close()

                st.session_state["excel_bytes_cache"] = updated_bytes
                st.session_state["updated_workbook"] = updated_bytes
                st.session_state["updated_sheet"] = selected_sheet
                st.session_state["update_message"] = (
                    f"Excel sheet '{selected_sheet}' updated successfully."
                )

            wb.close()
            workbook_buffer.close()

    if st.session_state.get("update_message"):
        st.success(st.session_state["update_message"])

    if st.session_state.get("updated_workbook"):
        st.download_button(
            "⬇️ Download updated workbook",
            data=st.session_state["updated_workbook"],
            file_name=f"updated_{excel.name}",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

else:
    st.warning("No valid data extracted")
