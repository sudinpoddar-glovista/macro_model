# IMPORT PYTHON LIBRARIES
import streamlit as st
import altair as alt
import pandas as pd


# EXTRACT MOST RECENT COUNTRY MODEL RANK CALCULATION EXCEL FILE
@st.experimental_memo
def get_fin_con_model_data():
    """
    Takes in the latest excel output from country model calculations
    :return: dataframes for each sheet in excel file
    """
    #folder_path = "I:/Sudin Poddar/Tools/PyCharm/ilc-main/output/ilc_macro_model/"
    #file_type = 'macro_model_data.xlsx'
    model_file = "macro_model_data.xlsx"

    # ASSIGN EACH SHEET OF EXCEL FILE TO A DATAFRAME
    xl = pd.ExcelFile(model_file)
    df_dict = {}
    for sheet in xl.sheet_names:
        df_dict[f'{sheet}'] = pd.read_excel(xl, sheet_name=sheet)

    return df_dict['FinCondition'], df_dict['IndexReturns']


# SELECT RECENT AND PREVIOUS DATES
def fin_con():
    # CALL THE MODEL DATA FUNCTION AND ASSIGN TO DF VARIABLES
    fin_con_df, idx_ret_df = get_fin_con_model_data()

    fin_con_df.columns = ['Date', 'Index', 'Value']
    idx_ret_df.columns = ['Date', 'Index', 'Return', 'ReturnType']

    # ASSIGN HEADER TO PAGE
    st.title("Chicago Fed National Financial Condition Data")

    # CREATE DROPDOWN TO SELECT RETURN STREAM
    st.caption("Return Stream Selection:")
    index_select = st.selectbox(
        'Select index:', idx_ret_df['Index'].unique(), key='index_select')
    return_select = st.selectbox(
        'Select return type:', idx_ret_df['ReturnType'].unique(), index=1, key='return_select')

    selected_idx_ret_df = idx_ret_df.loc[(idx_ret_df.Index == index_select) & (idx_ret_df.ReturnType == return_select)]
    selected_idx_ret_df.drop(columns=['ReturnType'], inplace=True)
    selected_idx_ret_df.columns = ['Date', 'Index', 'Return']
    selected_idx_ret_df['Return'] = selected_idx_ret_df['Return'].mul(100)

    min_dt = selected_idx_ret_df.at[selected_idx_ret_df['Return'].notna().idxmax(), 'Date']
    selected_idx_ret_df = selected_idx_ret_df.loc[selected_idx_ret_df.Date >= min_dt]
    selected_idx_ret_df.Return = selected_idx_ret_df.Return.fillna(0)

    selected_fin_cond_df = fin_con_df.loc[fin_con_df.Date >= min_dt]

    # CREATE CHART
    st.caption("Condition vs Returns Charts:")

    scales = alt.selection_interval(bind='scales')

    cond_chart = alt.Chart(selected_fin_cond_df).mark_bar().encode(
        x='Date:T',
        y='Value:Q',
        color=alt.condition(
            alt.datum.Value > 0,
            alt.value("#DB3951"),
            alt.value("#00136C")),
        tooltip=['Date', 'Index', 'Value']
    ).add_selection(scales).configure_mark(opacity=0.4)

    st.altair_chart(cond_chart, use_container_width=True)

    ret_chart = alt.Chart(selected_idx_ret_df).mark_bar().encode(
        x='Date:T',
        y=alt.Y('Return:Q', axis=alt.Axis(title='Return %')),
        color=alt.condition(
            alt.datum.Return > 0,
            alt.value("#6F9D80"),
            alt.value("#FA7711")),
        tooltip=['Date', 'Index', 'Return']
    ).add_selection(scales).configure_mark(opacity=0.4)

    st.altair_chart(ret_chart, use_container_width=True)


fin_con()
