# IMPORT PYTHON LIBRARIES
import streamlit as st
from st_aggrid import AgGrid
from st_aggrid.shared import JsCode
from st_aggrid.grid_options_builder import GridOptionsBuilder
import altair as alt
import pandas as pd


# EXTRACT MOST RECENT COUNTRY MODEL RANK CALCULATION EXCEL FILE
@st.experimental_memo
def get_model_data():
    """
    Takes in the latest excel output from country model calculations
    :return: dataframes for each sheet in excel file
    """
    #folder_path = "https://raw.githubusercontent.com/sudinrmg/macro_model/"
    file_ = 'country_model_data.xlsx'
    #model_file = f"{folder_path}{file_type}"

    # ASSIGN EACH SHEET OF EXCEL FILE TO A DATAFRAME
    df_dictionary = pd.read_excel(file_, sheet_name=None)
    return df_dictionary


# SELECT RECENT AND PREVIOUS DATES
def main():
    # CALL THE MODEL DATA FUNCTION AND ASSIGN TO DF VARIABLES
    df_dict = get_model_data()
    df_metric_score, df_factor_score, df_variable = df_dict.values()

    df_overall = df_metric_score.pivot(index=['Periodid', 'CountryName'], columns=['Metric'], values='Score')
    df_overall.reset_index(inplace=True)
    df_overall = df_overall[['Periodid', 'CountryName', 'Overall Rank', 'Overall', 'Economics',
                             'Valuation', 'Sentiment', 'Profit', 'Leverage']]
    df_overall.sort_values(by=['Periodid'], ascending=False, inplace=True)

    # FIX PERIODID DTYPE
    df_metric_score['Periodid'] = pd.to_datetime(df_metric_score['Periodid'])
    df_metric_score.set_index('Periodid', inplace=True)
    df_metric_score.index = df_metric_score.index.to_period('M').to_timestamp('M')
    df_metric_score.reset_index(inplace=True)

    df_factor_score['Periodid'] = pd.to_datetime(df_factor_score['Periodid'])
    df_factor_score.set_index('Periodid', inplace=True)
    df_factor_score.index = df_factor_score.index.to_period('M').to_timestamp('M')
    df_factor_score.reset_index(inplace=True)

    df_variable['Periodid'] = pd.to_datetime(df_variable['Periodid'])
    df_variable.set_index('Periodid', inplace=True)
    df_variable.index = df_variable.index.to_period('M').to_timestamp('M')
    df_variable.reset_index(inplace=True)

    # ASSIGN HEADER TO MAIN PAGE
    st.title("ILC Country Ranking Framework")

    tab1, tab2, tab3 = st.tabs(["Score Compare", "Score History", "Factor History"])

    # ASSIGN SUBHEADER AND CREATE SELECTION BOXES FOR DATES IN SIDEBAR
    with tab1:
        st.caption("Dates For Rank Compare:")
        recent_dt = st.selectbox(
            'Select recent date:', df_overall['Periodid'].unique(), key='recent_dt')
        prior_dt = st.selectbox(
            'Select prior date:', df_overall['Periodid'].unique(), index=1, key='prior_dt')

        recent_rnk_df = df_overall.loc[df_overall['Periodid'] == recent_dt]
        prior_rnk_df = df_overall.loc[df_overall['Periodid'] == prior_dt]
        rnk_comp_df = recent_rnk_df.merge(prior_rnk_df, how='left', on='CountryName', suffixes=['_r', '_p'])
        rnk_comp_df['Rnk_d'] = rnk_comp_df['Overall Rank_r'] - rnk_comp_df['Overall Rank_p']
        rnk_comp_df['Econ_d'] = rnk_comp_df['Economics_r'] - rnk_comp_df['Economics_p']
        rnk_comp_df['Val_d'] = rnk_comp_df['Valuation_r'] - rnk_comp_df['Valuation_p']
        rnk_comp_df['Sent_d'] = rnk_comp_df['Sentiment_r'] - rnk_comp_df['Sentiment_p']
        rnk_comp_df['Prof_d'] = rnk_comp_df['Profit_r'] - rnk_comp_df['Profit_p']
        rnk_comp_df['Lev_d'] = rnk_comp_df['Leverage_r'] - rnk_comp_df['Leverage_p']
        rnk_comp_df = rnk_comp_df[['Periodid_r', 'CountryName',
                                   'Overall Rank_r', 'Overall Rank_p', 'Rnk_d',
                                   'Economics_r', 'Economics_p', 'Econ_d',
                                   'Valuation_r', 'Valuation_p', 'Val_d',
                                   'Sentiment_r', 'Sentiment_p', 'Sent_d',
                                   'Profit_r', 'Profit_p', 'Prof_d',
                                   'Leverage_r', 'Leverage_p', 'Lev_d'
                                   ]]
        col_rename = {'Periodid_r': 'Date', 'CountryName': 'Country',
                      'Overall Rank_r': 'Rnk_r', 'Overall Rank_p': 'Rnk_p',
                      'Economics_r': 'Econ_r', 'Economics_p': 'Econ_p',
                      'Valuation_r': 'Val_r', 'Valuation_p': 'Val_p',
                      'Sentiment_r': 'Sent_r', 'Sentiment_p': 'Sent_p',
                      'Profit_r': 'Prof_r', 'Profit_p': 'Prof_p',
                      'Leverage_r': 'Lev_r', 'Leverage_p': 'Lev_p'
                      }
        rnk_comp_df.rename(columns=col_rename, inplace=True)
        rnk_comp_df = rnk_comp_df.round(2)
        rnk_comp_df.sort_values(by=['Rnk_r'], ascending=False, inplace=True)

        # ASSIGN SUBHEADER AND BUILD FORMATTED TABLE
        st.caption("Country Rank Delta Table:")

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

        gb = GridOptionsBuilder.from_dataframe(rnk_comp_df, minWidth=10)
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

        @st.cache_data
        def convert_df(df):
            # IMPORTANT: Cache the conversion to prevent computation on every rerun
            return df.to_csv().encode('utf-8')

        csv = convert_df(rnk_comp_df)

        st.download_button(
            label="Download data as CSV",
            data=csv,
            file_name='rnk_comp_df.csv',
            mime='text/csv',
        )
        
    with tab2:

        # ASSIGN SUBHEADER AND CREATE SELECTION BOXES
        st.caption("Filters For Score History:")

        metrics_unique = ['Overall', 'Economics', 'Valuation', 'Sentiment', 'Profit', 'Leverage']
        colors_unique = ['#3cb44b', '#469990', '#4363d8', '#aaffc3', '#dcbeff', '#ffd8b1']

        cntry_nm = st.multiselect(
            'Select country/s:', df_metric_score['CountryName'].unique(), ['China'], key='sh_cntry')
        metric_nm = st.multiselect(
            'Select metric/s:', metrics_unique, ['Overall'], key='sh_metric')

        # CREATE THE METRIC HISTORY TABLE
        score_hist_df = df_metric_score.loc[(df_metric_score['CountryName'].isin(cntry_nm)) &
                                            (df_metric_score['Metric'].isin(metric_nm))]

        # ASSIGN SUBHEADER AND CALL CHART
        st.caption("Score History Chart:")

        scales = alt.selection_interval(bind='scales')
        score_hist_chart = alt.Chart(score_hist_df).mark_line().encode(
            x='Periodid:T',
            y='Score:Q',
            color=alt.Color('Metric', scale=alt.Scale(domain=metrics_unique, range=colors_unique)),
            strokeDash='CountryName:N',
            tooltip=['Periodid', 'CountryName', 'Metric', 'Score']
        ).add_selection(scales)

        st.altair_chart(score_hist_chart, use_container_width=True)

    with tab3:

        # ASSIGN SUBHEADER AND CREATE SELECTION BOXES
        st.caption("Filters For Factor History:")

        metric_fh = st.selectbox(
            'Select metric:', df_variable['Metric'].unique(), key='fh_metric')
        factor_fh = st.selectbox(
            'Select factor:', df_variable.loc[df_variable['Metric'] == metric_fh]['Factor'].unique(),
            key='fh_factor')
        country_fh = st.multiselect(
            'Select country/s:', df_variable['CountryName'].unique(), cntry_nm, key='fh_cntry')

        # CREATE THE FACTOR HISTORY TABLE
        df_variable_hist = df_variable.loc[(df_variable['Factor'] == factor_fh) &
                                           (df_variable['Periodid'] >= df_factor_score['Periodid'].min()) &
                                           (df_variable['CountryName'].isin(country_fh))][['Periodid',
                                                                                           'CountryName',
                                                                                           'value_orig']]

        variable_score_df = df_factor_score.loc[(df_factor_score['Factor'] == factor_fh) &
                                                (df_factor_score['CountryName'].isin(country_fh))][['Periodid',
                                                                                                    'CountryName',
                                                                                                    'FactorZScore']]

        # CREATE SIDE BY SIDE CHARTS
        col1, col2 = st.columns(2)

        with col1:
            st.caption("Factor History Chart:")

            fact_hist_chart = alt.Chart(df_variable_hist).mark_line().encode(
                x='Periodid:T',
                y='value_orig:Q',
                color=alt.Color('CountryName:N', scale=alt.Scale(scheme='darkmulti')),
                tooltip=['Periodid', 'CountryName', 'value_orig']
            ).add_selection(scales)

            st.altair_chart(fact_hist_chart, use_container_width=True)

        with col2:
            st.caption("Factor Score History Chart:")

            fact_score_chart = alt.Chart(variable_score_df).mark_line().encode(
                x='Periodid:T',
                y='FactorZScore:Q',
                color=alt.Color('CountryName:N', scale=alt.Scale(scheme='darkmulti')),
                tooltip=['Periodid', 'CountryName', 'FactorZScore']
            ).add_selection(scales)

            st.altair_chart(fact_score_chart, use_container_width=True)


#

# factor_hist_df = factor_hist_df.fillna(0)
# factor_hist_df['Periodid'] = pd.to_datetime(factor_hist_df['Periodid'])
# factor_hist_df.set_index('Periodid', inplace=True)
# factor_hist_df.index = factor_hist_df.index.to_period('M').to_timestamp('M')
# factor_hist_df.reset_index(inplace=True)
#
# # CALL HEADER AND CHART
# st.header("Country Factor History:")
# factor_chart = alt.Chart(factor_hist_df).mark_line().encode(
#     x='Periodid', y='value_orig', color='CountryName', tooltip=['Periodid', 'CountryName',
#                                                                 'value_orig'])
#
# st.altair_chart(factor_chart, use_container_width=True)


## RUN THE APP IF CALLED FROM THIS SCRIPT
if __name__ == "__main__":
    st.set_page_config("ILC Country Ranking Framework",
        #        initial_sidebar_state="expanded", 
                       layout="wide")
main()
