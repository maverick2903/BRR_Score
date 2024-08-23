import streamlit as st
import plotly.express as px
import pandas as pd
import os

# Define the scoring function based on the given criteria
def calculate_score(row, yes_no_columns, range_columns, list_columns):
    score = 0

    for col in yes_no_columns:
        if row[col] == 'Yes':
            score += 1

    # Normalize and compute ranking scores
    for col in range_columns:
        if pd.notna(row[col]):
            rank = int(row[col])
            normalized_rank = (rank - 1) / (4 - 1)
            score += normalized_rank

    # Normalize and compute list-based scores
    for col in list_columns:
        if pd.notna(row[col]):
            items = row[col].split(',')  # Assuming items are comma-separated
            normalized_list_score = len(items) / 9
            score += normalized_list_score

    return score

# Streamlit app
st.title("BRR Score Calculator")

# File upload
uploaded_file = st.file_uploader("Upload Excel file", type=["xlsx"])

if uploaded_file is not None:
    file_name = os.path.splitext(uploaded_file.name)[0]
    # Load the Excel file
    sheet_name = st.text_input("Enter sheet name to be parsed: ", "Sheet4")
    df = pd.read_excel(uploaded_file, sheet_name=sheet_name)
    
    # Display the first few rows of the data
    st.write("Data Preview:")
    st.write(df.head())

    # Drop the unnecessary column
    if 'Unnamed: 0' in df.columns:
        df.drop(['Unnamed: 0'], axis=1, inplace=True)

    if 'Timestamp' in df.columns:
        df.drop(['Timestamp'], axis=1, inplace=True)

    # User inputs for row index
    # row_index = st.number_input("Enter the row index for sorting column names:", min_value=0, max_value=len(df)-1, value=1)
    row_index = 1

    # Define the criteria for sorting column names
    criteria1 = lambda x: x == 'Yes' or x == 'No' or x == 'Not clear/Not mentioned' or x == 'Not clear' or x == 'Not clear/Not mentioned/NA'
    criteria2 = lambda x: x == 1 or x == 2 or x == 3 or x == 4
    criteria3 = lambda x: isinstance(x, str) and 'P1' in x

    # Get the values in the specified row
    row_values = df.iloc[row_index]

    # Sort column names into three lists based on the criteria
    columns_criteria1 = [col for col in df.columns if criteria1(row_values[col])]
    columns_criteria2 = [col for col in df.columns if criteria2(row_values[col])]
    columns_criteria3 = [col for col in df.columns if criteria3(row_values[col])]

    # Display the sorted column names
    st.write("Columns matching criteria 1 (YES/NO):", columns_criteria1)
    st.write("Columns matching criteria 2 (1,2,3,4):", columns_criteria2)
    st.write("Columns matching criteria 3 (P1,P2,P3...):", columns_criteria3)

    # Apply the scoring function to each row in the dataframe
    df['Calculated Score'] = df.apply(calculate_score, axis=1, yes_no_columns=columns_criteria1, range_columns=columns_criteria2, list_columns=columns_criteria3)

    # Calculate the total possible score
    total_possible_score = len(columns_criteria1) + len(columns_criteria2) + len(columns_criteria3)
    st.write(f"Total possible score: {total_possible_score}")

    # Display the dataframe with the new calculated scores and the total possible score
    df1 = df[['NAME OF THE COMPANY', 'YEAR OF REPORTING', 'Calculated Score', 'SECTOR']].dropna()
    df1['Percentage Score'] = df1['Calculated Score'] / total_possible_score
    st.write("Calculated Scores:")
    st.write(df1)

    # Provide an option to download the calculated scores
    if st.button("Download Scores"):
        output_file_name = f"{file_name}_{sheet_name}.xlsx"
        df1.to_excel(output_file_name, index=False)
        with open(output_file_name, "rb") as file:
            st.download_button(
                label="Download Excel file",
                data=file,
                file_name=output_file_name,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

    # Show a bar chart of the calculated scores by sector
    st.write("Bar Chart: Calculated Scores by Sector")
    st.bar_chart(df1.groupby('SECTOR')['Calculated Score'].mean())

    # Show a line chart of the percentage scores over time
    st.write("Line Chart: Percentage Scores over Time")
    st.line_chart(df1.groupby('YEAR OF REPORTING')['Percentage Score'].mean())

    st.write("Histogram: Distribution of Calculated Scores")
    fig = px.histogram(df1, x='Calculated Score', nbins=20, title="Distribution of Calculated Scores")
    st.plotly_chart(fig)

    st.write("Box Plot: Score Comparison by Sector")
    fig = px.box(df1, x='SECTOR', y='Calculated Score', title="Score Comparison by Sector")
    st.plotly_chart(fig)

    st.write("Line Chart: Score Trends by Sector Over Time")
    fig = px.line(df1, x='YEAR OF REPORTING', y='Percentage Score', color='SECTOR', title="Score Trends by Sector Over Time")
    st.plotly_chart(fig)

    st.write("Top Performers")
    top_performers = df1.nlargest(5, 'Calculated Score')
    st.write(top_performers)

    st.write("Bottom Performers")
    bottom_performers = df1.nsmallest(5, 'Calculated Score')
    st.write(bottom_performers)

    st.write("Violin Plot: Sector-wise Score Distribution")
    fig = px.violin(df1, x='SECTOR', y='Calculated Score', box=True, points='all', title="Sector-wise Score Distribution")
    st.plotly_chart(fig)

    # st.write("Correlation Heatmap")
    # numeric_columns = df1.select_dtypes(include=['float64', 'int64']).columns
    # corr = df1[numeric_columns].corr()
    # fig = px.imshow(corr, text_auto=True, title="Correlation Heatmap")
    # st.plotly_chart(fig)

    st.write("Pie Chart: Sector Representation")
    fig = px.pie(df1, names='SECTOR', title="Sector Representation in Dataset")
    st.plotly_chart(fig)

    st.write("Year-over-Year Score Improvement")
    df1['Previous Year Score'] = df1.groupby('NAME OF THE COMPANY')['Percentage Score'].shift(1)
    df1['YoY Improvement'] = df1['Percentage Score'] - df1['Previous Year Score']
    fig = px.bar(df1.dropna(), x='YEAR OF REPORTING', y='YoY Improvement', color='SECTOR', title="Year-over-Year Score Improvement")
    st.plotly_chart(fig)

    st.write("Filter Data")
    selected_sector = st.selectbox("Select Sector", df1['SECTOR'].unique())
    filtered_df = df1[df1['SECTOR'] == selected_sector]
    st.write(filtered_df)

    if st.button("Download Filtered Data"):
        filtered_output_file_name = f"{file_name}_{sheet_name}_{selected_sector}.xlsx"
        filtered_df.to_excel(filtered_output_file_name, index=False)
        with open(filtered_output_file_name, "rb") as file:
            st.download_button(
                label="Download Filtered Excel file",
                data=file,
                file_name=filtered_output_file_name,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

