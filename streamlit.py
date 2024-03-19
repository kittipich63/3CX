import os
import re
import streamlit as st
import pandas as pd
import plotly.express as px
from io import BytesIO
import subprocess

def install_requirements(requirements):
    """
    Install Python packages listed in the requirements string.
    """
    packages = re.findall(r'\w+==[\d.]+', requirements)
    for package in packages:
        subprocess.call(['pip', 'install', package])


def process_data(data):
    # Create separate columns for USE and NO USE based on Status
    data["USE"] = data["Status"].apply(lambda x: 1 if x == "USE" else 0)
    data["NO USE"] = data["Status"].apply(lambda x: 1 if x == "NO USE" else 0)

    # Group by Location and sum the USE and NO USE columns
    grouped_data = (
        data.groupby("Location")
        .agg({"NO USE": "sum", "USE": "sum"})
        .reset_index()
    )

    # Add Register to server column and calculate the sum for each row
    grouped_data["Register to server"] = grouped_data["USE"] + grouped_data["NO USE"]

    # Add a row named "Total" by counting the number of rows in each column
    total_row = pd.DataFrame({
        'Location': ['Total'],
        'USE': [grouped_data['USE'].sum()],
        'NO USE': [grouped_data['NO USE'].sum()],
        'Register to server': [grouped_data['Register to server'].sum()]
    })

    grouped_data = pd.concat([grouped_data,total_row], ignore_index=True)

    return grouped_data
        
def main():
    st.set_page_config(layout="wide")
    st.title(f"Summary Report User Used(Factory)  Not Used 3CX ")

    # Install requirements
    requirements = """
    pandas==2.1.4
    streamlit==1.29.0
    plotly==5.18.0
    plotly-express==0.4.1
    XlsxWriter==3.1.9
    """
    install_requirements(requirements)

    # Upload CSV file
    uploaded_file = st.file_uploader("Upload File", type=["csv"])

    if uploaded_file is not None:

        # Get the name of the uploaded file without the extension
        file_name = os.path.splitext(uploaded_file.name)[0]

        # Read CSV file
        data = pd.read_csv(uploaded_file)

        # Process data based on conditions
        processed_data = process_data(data)

###################################  Display processed data  ###################################
        st.subheader(f"Report User Used  /  Not Used  of {file_name}")
        st.dataframe(processed_data, width=None)

###################################  Create a new DataFrame for plotting the histogram  ###################################
        hist_data = pd.DataFrame({
            "Location": processed_data["Location"],
            "USE": processed_data["USE"],
            "NO USE": processed_data["NO USE"],
            # "Register to server": processed_data["Register to server"]
        })

        # Exclude the "Total" row from the histogram plot
        hist_data = hist_data[hist_data["Location"] != "Total"]

###################################  Create a histogram plot for USE and NO USE data  ###################################
        st.subheader("Histogram of USE and NO USE Data by Location")
        hist_fig = px.histogram(hist_data, y="Location", x=["USE", "NO USE"], barmode="group")
        st.plotly_chart(hist_fig)


###################################  Show Section From Location  ###################################
        # Create a dropdown for selecting the location
        selected_location_section = st.selectbox(
            "Select Location (Section):", processed_data["Location"].loc[processed_data["Location"] != "Total"].tolist()
        )

        # Display detailed information for the selected location without USE and NO USE columns
        st.subheader(f"Detailed Information Section for  {selected_location_section}")
        selected_data_section = data[data["Location"] == selected_location_section].groupby("Section").agg({
            "Call": "sum",
            "Receive": "sum",
            "Total": "sum"
        }).reset_index()
        st.dataframe(selected_data_section, use_container_width=True)

        # Export data to CSV file all Location
        export_button_key_section = f"export_button_{selected_data_section}"
        if st.button("Export to Excel", key=export_button_key_section):
            # Write to a single Excel file with sheets divided by Location
            file_path = "data_section_location.xlsx"
            with pd.ExcelWriter(file_path, engine="xlsxwriter") as writer:
                unique_locations = data["Location"].unique()
                for location_name in unique_locations:
                    # Clean the location name to remove invalid characters
                    cleaned_location_name = re.sub(r'[^\w\s]', '', location_name)

                    location_data = data[data["Location"] == location_name].groupby("Section").agg({
                        "Call": "sum",
                        "Receive": "sum",
                        "Total": "sum"
                    }).reset_index()
                    location_data.to_excel(writer, sheet_name=cleaned_location_name, index=False)

            # Provide a download button for the user
            st.download_button(
                label="Download Data",
                data=open(file_path, 'rb').read(),
                key="download_data",
                file_name="Section From Location (All).xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )


###################################  Show Information From Location  ###################################
        #Create a dropdown for selecting the location
        selected_location = st.selectbox(
            "Select Location (User) :", processed_data["Location"].loc[processed_data["Location"] != "Total"].tolist()
        )

        #Display detailed information for the selected location without USE and NO USE columns
        st.subheader(f"Detailed Information User for {selected_location}")        
        selected_data = data[data["Location"] == selected_location].drop(columns=["USE", "NO USE"])

        # Format the "Number" column to remove commas
        selected_data["Number"] = selected_data["Number"].apply(lambda x: '{:.0f}'.format(x))

        # Display the dataframe
        st.dataframe(selected_data, use_container_width=True)

        # Button to trigger export to Excel
        export_button_key_user = f"export_button_{selected_location}"
        if st.button("Export to Excel", key=export_button_key_user):
            # Export data to Excel file with sanitized sheets
            excel_file_path = "data_user_location.xlsx"
            with pd.ExcelWriter(excel_file_path, engine="xlsxwriter") as writer:
                for location_name, location_data in data.groupby("Location"):
                    if location_name != "Total":
                        # Sanitize the sheet name by removing invalid characters
                        sanitized_location_name = re.sub(r'[^\w\s]', '', location_name)
                        location_data.drop(columns=["USE", "NO USE"]).to_excel(writer, sheet_name=sanitized_location_name, index=False)

            # Provide a download button for the user
            st.download_button(
                label="Download Data",
                data=open(excel_file_path, 'rb').read(),
                key="download_data",
                file_name="Location By (All).xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )


if __name__ == '__main__':
    main()