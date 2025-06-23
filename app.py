import streamlit as st
import requests
import pandas as pd
from io import BytesIO
from openpyxl.utils import get_column_letter

# Replace with your actual API key
EASYPOST_API_KEY = "EZTK1bdc79ebc5044ca8a44cd56fa7c34d0eakqM9V1LNtDI6RfbpA49wQ"

# Function to validate address
def validate_address_easypost(address):
    url = "https://api.easypost.com/v2/addresses/verify"
    headers = {"Authorization": f"Bearer {EASYPOST_API_KEY}"}
    payload = {
        "address": {
            "street1": address["Address Line 1"],
            "street2": address.get("Address Line 2", ""),
            "city": address["City"],
            "zip": address["Postal Code"],
            "country": address["Country"]
        }
    }

    response = requests.post(url, json=payload, headers=headers)
    response_json = response.json()

    if "address" in response_json:
        verifications = response_json["address"].get("verifications", {})
        delivery_status = verifications.get("delivery", {}).get("success", False)
        confidence = "High Confidence" if delivery_status else "Low Confidence !!!NEED CHECK!!!"
    else:
        confidence = "Low Confidence !!!NEED CHECK!!!"

    return response_json, confidence

# Format the verified address
def format_address_easypost(response):
    if "address" in response:
        addr = response["address"]
        formatted = f"{addr.get('street1', '')}, {addr.get('street2', '')}, {addr.get('city', '')}, {addr.get('state', '')} {addr.get('zip', '')}, {addr.get('country', '')}"
        return formatted.strip(", ")
    else:
        return "Invalid address or could not be verified"

# --- Streamlit UI ---
st.set_page_config(
    page_title="EasyPost Address Validator",
    page_icon="📦",
    layout="wide"
)
st.markdown("### 📦 EasyPost Address Validator")

uploaded_file = st.file_uploader("Upload an Excel file with addresses", type=["xlsx"])

if uploaded_file:
    try:
        df = pd.read_excel(uploaded_file)
        df.fillna("", inplace=True)
        df = df.map(lambda x: x.upper() if isinstance(x, str) else x)

        # Validate required columns
        required_columns = ["Address Line 1", "City", "Postal Code", "Country"]
        if not all(col in df.columns for col in required_columns):
            st.error("Missing one or more required columns: Address Line 1, City, Postal Code, Country")
        else:
            results = []
            with st.spinner("Validating addresses..."):
                for _, row in df.iterrows():
                    address = {
                        "Address Line 1": row.get("Address Line 1", ""),
                        "Address Line 2": row.get("Address Line 2", ""),
                        "City": row.get("City", ""),
                        "Postal Code": str(row.get("Postal Code", "")),
                        "Country": row.get("Country", "")
                    }

                    try:
                        response, confidence = validate_address_easypost(address)
                        formatted = format_address_easypost(response)
                    except Exception as e:
                        formatted = f"ERROR: {str(e)}"
                        confidence = "Error"

                    results.append({
                        "Original Address": f"{address['Address Line 1']}, {address['City']}, {address['Postal Code']}, {address['Country']}",
                        "Formatted Address": formatted,
                        "Confidence": confidence
                    })

            result_df = pd.DataFrame(results)
            st.success("Validation completed!")
            st.dataframe(result_df)

            # Optional: allow download


            output = BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                result_df.to_excel(writer, index=False, sheet_name='Validated Results')

                # Auto-adjust column widths
                worksheet = writer.sheets['Validated Results']
                for i, column in enumerate(result_df.columns, start=1):
                    max_length = max(
                        result_df[column].astype(str).map(len).max(),  # max length in column
                        len(column)  # header length
                    )
                    worksheet.column_dimensions[get_column_letter(i)].width = max_length + 2  # padding

            output.seek(0)

            st.download_button(
                label="Download Results as Excel",
                data=output,
                file_name="validated_addresses.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

            
            # st.download_button("Download Results as Excel", data=result_df.to_excel(index=False), file_name="validated_addresses.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    except Exception as e:
        st.error(f"Failed to process file: {str(e)}")
