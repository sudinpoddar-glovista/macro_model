# IMPORT PYTHON LIBRARIES
import streamlit as st
from st_aggrid import AgGrid
from st_aggrid.shared import JsCode
from st_aggrid.grid_options_builder import GridOptionsBuilder
import altair as alt
import pandas as pd
import glob
import os.path


# EXTRACT MOST RECENT COUNTRY MODEL RANK CALCULATION EXCEL FILE
@st.experimental_memo
def get_model_data():
    """
    Takes in the latest excel output from country model calculations
    :return: dataframes for each sheet in excel file
    """
    folder_path = r'I:\Sudin Poddar\Tools\PyCharm\ilc-main\output\ilc_country_model'
    file_type = r'\*xlsx'
    files = glob.glob(folder_path + file_type)
    max_file = max(files, key=os.path.getctime)

    # ASSIGN EACH SHEET OF EXCEL FILE TO A DATAFRAME
    xl = pd.ExcelFile(max_file)
    df_dict = {}
    for sheet in xl.sheet_names:
        df_dict[f'{sheet}'] = pd.read_excel(xl, sheet_name=sheet)

    final_rnk = df_dict['FinalRank']
    metric_rnk = df_dict['MetricRank']
    factor_rnk = df_dict['FactorRank']
    variable_rnk = df_dict['VariableRank']
    variable = df_dict['Variable']

    # CREATE A COMBINED DATAFRAME FOR FINAL AND METRIC RANKS
    overall_rnk = final_rnk.merge(metric_rnk, how='left', on=['Periodid', 'CountryName'])
    overall_rnk = overall_rnk[['Periodid', 'CountryName', 'CountryPctRnk', 'CountryZScore',
                               'Economics', 'Valuation', 'Sentiment', 'Profit', 'Leverage']]
    overall_rnk.rename(columns={'CountryZScore': 'Overall'}, inplace=True)
    return final_rnk, metric_rnk, factor_rnk, variable_rnk, variable, overall_rnk


# SELECT RECENT AND PREVIOUS DATES
def main():
    # CALL THE MODEL DATA FUNCTION AND ASSIGN TO DF VARIABLES
    final_rnk_df, metric_rnk_df, factor_rnk_df, variable_rnk_df, variable_df, overall_rnk_df = get_model_data()

    # ASSIGN HEADER TO MAIN PAGE
    st.title("ILC Country Ranking Framework")

    # ASSIGN SUBHEADER AND CREATE SELECTION BOXES FOR DATES IN SIDEBAR
    st.sidebar.subheader("Dates For Rank Compare:")
    recent_dt = st.sidebar.selectbox(
        'Select recent date:', overall_rnk_df['Periodid'].unique(), key='recent_dt')
    prior_dt = st.sidebar.selectbox(
        'Select prior date:', overall_rnk_df['Periodid'].unique(), index=1, key='prior_dt')

    # CREATE THE RANK COMPARE TABLE
    recent_rnk_df = overall_rnk_df.loc[overall_rnk_df['Periodid'] == recent_dt]
    prior_rnk_df = overall_rnk_df.loc[overall_rnk_df['Periodid'] == prior_dt]
    rnk_comp_df = recent_rnk_df.merge(prior_rnk_df, how='left', on='CountryName', suffixes=['_r', '_p'])
    rnk_comp_df['Rnk_d'] = rnk_comp_df['CountryPctRnk_r'] - rnk_comp_df['CountryPctRnk_p']
    rnk_comp_df['Econ_d'] = rnk_comp_df['Economics_r'] - rnk_comp_df['Economics_p']
    rnk_comp_df['Val_d'] = rnk_comp_df['Valuation_r'] - rnk_comp_df['Valuation_p']
    rnk_comp_df['Sent_d'] = rnk_comp_df['Sentiment_r'] - rnk_comp_df['Sentiment_p']
    rnk_comp_df['Prof_d'] = rnk_comp_df['Profit_r'] - rnk_comp_df['Profit_p']
    rnk_comp_df['Lev_d'] = rnk_comp_df['Leverage_r'] - rnk_comp_df['Leverage_p']
    rnk_comp_df = rnk_comp_df[['Periodid_r', 'CountryName',
                               'CountryPctRnk_r', 'CountryPctRnk_p', 'Rnk_d',
                               'Economics_r', 'Economics_p', 'Econ_d',
                               'Valuation_r', 'Valuation_p', 'Val_d',
                               'Sentiment_r', 'Sentiment_p', 'Sent_d',
                               'Profit_r', 'Profit_p', 'Prof_d',
                               'Leverage_r', 'Leverage_p', 'Lev_d'
                               ]]
    col_rename = {'Periodid_r': 'Date', 'CountryName': 'Country',
                  'CountryPctRnk_r': 'Rnk_r', 'CountryPctRnk_p': 'Rnk_p',
                  'Economics_r': 'Econ_r', 'Economics_p': 'Econ_p',
                  'Valuation_r': 'Val_r', 'Valuation_p': 'Val_p',
                  'Sentiment_r': 'Sent_r', 'Sentiment_p': 'Sent_p',
                  'Profit_r': 'Prof_r', 'Profit_p': 'Prof_p',
                  'Leverage_r': 'Lev_r', 'Leverage_p': 'Lev_p'
                  }
    rnk_comp_df.rename(columns=col_rename, inplace=True)
    rnk_comp_df = rnk_comp_df.round(2)

    # ASSIGN SUBHEADER AND BUILD FORMATTED TABLE
    st.header("Country Rank Delta Table:")

    # JSCON CELL STYLE
    cellsytle_jscode = JsCode(
        """
    function(params) {
        if (params.value > 0) {
            return {
                'color': 'white',
                'backgroundColor': '#00136C'
            }
        } else if (params.value < 0) {
            return {
                'color': 'white',
                'backgroundColor': '#DB3951'
            }
        } else {
            return {
                'color': 'black',
                'backgroundColor': '#F0F3F3'
            }
        }
    };
    """
    )

    gb = GridOptionsBuilder.from_dataframe(rnk_comp_df, minWidth=20)
    gb.configure_columns(
        (
            "Rnk_d",
            "Econ_d",
            "Val_d",
            "Sent_d",
            "Prof_d",
            "Lev_d",
        ),
        cellStyle=cellsytle_jscode,
    )
    gb.configure_pagination(enabled=False)
    gb.configure_columns(("Date", "Country"), pinned=True)
    gb.configure_selection(selection_mode='multiple')
    gridoptions = gb.build()

    # CALL TABLE
    AgGrid(rnk_comp_df, gridOptions=gridoptions, allow_unsafe_jscode=True)

    # ASSIGN SUBHEADER AND CREATE SELECTION BOXES FOR DATES IN SIDEBAR
    st.sidebar.subheader("Score History:")
    cntry_nm = st.sidebar.selectbox(
        'Select country:', overall_rnk_df['CountryName'].unique(), key='single_cntry')
    metric_nm = st.sidebar.multiselect(
        'Select metric/s:', ['Overall', 'Economics', 'Valuation', 'Sentiment', 'Profit', 'Leverage'],
        ['Overall', 'Economics', 'Valuation'],
        key='multi_metric')

    # CREATE THE METRIC HISTORY TABLE
    country_hist_df = overall_rnk_df.loc[overall_rnk_df['CountryName'] == cntry_nm][['Periodid', *metric_nm]]
    country_hist_df['Periodid'] = pd.to_datetime(country_hist_df['Periodid'])
    country_hist_df.set_index('Periodid', inplace=True)
    country_hist_df.index = country_hist_df.index.to_period('M').to_timestamp('M')

    # ASSIGN SUBHEADER AND CALL CHART
    st.header("Country Metric Score History:")
    st.line_chart(country_hist_df)

    # ASSIGN SUBHEADER AND CREATE SELECTION BOXES FOR DATES IN SIDEBAR
    st.sidebar.subheader("Factor History:")
    metric_single = st.sidebar.selectbox(
        'Select metric:', variable_df['Metric'].unique(), key='single_metric')
    factor_single = st.sidebar.selectbox(
        'Select factor:', variable_df.loc[variable_df['Metric'] == metric_single]['Factor'].unique(),
        key='single_factor')
    country_multi = st.sidebar.multiselect(
        'Select country/s:', variable_df['CountryName'].unique(),
        cntry_nm,
        key='multi_cntry')

    # CREATE THE FACTOR HISTORY TABLE
    factor_hist_df = variable_df.loc[(variable_df['Factor'] == factor_single) &
                                     (variable_df['Periodid'] >= final_rnk_df['Periodid'].min()) &
                                     (variable_df['CountryName'].isin(country_multi))][['Periodid', 'CountryName',
                                                                                        'value_orig']]
    factor_hist_df = factor_hist_df.fillna(0)
    factor_hist_df['Periodid'] = pd.to_datetime(factor_hist_df['Periodid'])
    factor_hist_df.set_index('Periodid', inplace=True)
    factor_hist_df.index = factor_hist_df.index.to_period('M').to_timestamp('M')
    factor_hist_df.reset_index(inplace=True)

    # CALL HEADER AND CHART
    st.header("Country Factor History:")
    factor_chart = alt.Chart(factor_hist_df).mark_line().encode(
        x='Periodid', y='value_orig', color='CountryName', tooltip=['Periodid', 'CountryName',
                                                                                        'value_orig'])

    st.altair_chart(factor_chart, use_container_width=True)


# RUN THE APP IF CALLED FROM THIS SCRIPT
if __name__ == "__main__":
    st.set_page_config(
        "ILC Country Ranking Framework",
        initial_sidebar_state="expanded",
        layout="wide",
    )
    main()
