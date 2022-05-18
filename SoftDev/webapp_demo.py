from docxtpl import DocxTemplate
import docx
import pandas as pd
import streamlit as st
import streamlit_authenticator as stauth
from pathlib import Path
import base64
import os
import json
import pickle
import uuid
import re
from st_aggrid import AgGrid
from st_aggrid.grid_options_builder import GridOptionsBuilder

# be able to select each company
# figure out the score filtering 
# maybe add a folder in github with all the word docs
# have people choose the companies for them to download the word docs
# and only show the company name the score and the notes 

st.set_page_config(page_title='Post-Show Dashboard', page_icon=':bar_chart:', layout='wide')

names = ['demo_email']
usernames = ['demo_user']
passwords = ['demo_account']

hashed_passwords = stauth.Hasher(passwords).generate()

authenticator = stauth.Authenticate(names, usernames, hashed_passwords,
    'cOOkiE_poStSHowcHasINgAlL', 'keyY1969chasinGthEshoWsS', cookie_expiry_days=15)

name, authentication_status, username = authenticator.login('Login','main')

if authentication_status:
    # emojis: https://www.webfx.com/tools/emoji-cheat-sheet/

    # ---- MAINPAGE ----
    st.title(':bar_chart: Post-Show Dashboard')
    st.markdown('##')

    # ---- READ EXCEL ----
    @st.cache
    def get_data_from_excel(sheet):
        path_excel = Path(__file__).parents[1] / 'Mar22_Show/webapp_demo.xlsx' # demo file 
        df = pd.read_excel(
            io = path_excel,
            engine = 'openpyxl',
            sheet_name = sheet)
        df = df.astype(str)
        return df

    # add if user gets overall or only one of the categories or 2...
    # use second df for exports with all of the data
    dfshow = get_data_from_excel('TotalShow')
    dfex = get_data_from_excel('TotalEx')

    # ---- SIDEBAR ----
    st.sidebar.header('Please Filter Here:')

    new = st.sidebar.radio('Only companies new in this show?', ('Yes', 'No'))

    state = st.sidebar.multiselect('Select the State:',
        options=dfshow['State'].unique(),
        default=dfshow['State'].unique() )

    mobility_score = st.sidebar.multiselect('Select the Mobility score:',
        options=dfshow['mobility_ranking'].unique(),
        default=['1', '2', '3', '4'] )

    ucaas_score = st.sidebar.multiselect('Select the Ucaas/Ccaas score:',
        options=dfshow['ucaas_ccaas_ranking'].unique(),
        default=['1', '2', '3', '4'] )

    cyber_score = st.sidebar.multiselect('Select the Cyber score:',
        options=dfshow['cyber_ranking'].unique(),
        default=['1', '2', '3', '4'] )

    data_score = st.sidebar.multiselect('Select the Data Center score:',
        options=dfshow['DATA_Center_ranking'].unique(),
        default=['1', '2', '3', '4'] )
    
    if new == 'Yes':
        df_selection = dfex.query('(State == @state) | (mobility_ranking == @mobility_score) | (ucaas_ccaas_ranking == @ucaas_score) | (cyber_ranking == @cyber_score) | (DATA_Center_ranking == @data_score)')
    else:
        df_selection = dfshow.query('(State == @state) | (mobility_ranking == @mobility_score) | (ucaas_ccaas_ranking == @ucaas_score) | (cyber_ranking == @cyber_score) | (DATA_Center_ranking == @data_score)')
    
    st.dataframe(df_selection)

    selected_indices = st.multiselect('Select rows:', df_selection.index)
    selected_rows = df_selection.loc[selected_indices]
    st.write('### Selected Rows', selected_rows)

    # CSV Download button 
    st.download_button(label = 'Export current selection to CSV', data = selected_rows.to_csv(), mime='text/csv')

    # https://discuss.streamlit.io/t/how-to-take-text-input-from-a-user/187/3

# replace pdf with word, and add direct linking for some better one or 2 pagers. 
# use https://automatetheboringstuff.com/chapter13/ python-docx
    keepcols = ['Company',
    'Job Title',
    'State',
    'Department Spend',
    'Attendee Location',
    'Industry Sector',
    'Key Products or Services',
    'Employee Count',
    'Annual Sales',
    'Locations ',
    'IT Department Size',
    'IT Security Team Size',
    'Contact Center Seats',
    'Operating System',
    'Current ERP',
    'Cloud Service Provider']


# transfers the variables in the df to word doc
    def to_docs(company,df1):
        df = df1[keepcols]
        to_docx = df.loc[df['Company'] == company]
        compani = to_docx['Company'].iloc[0]
        state = to_docx['State'].iloc[0]
        job_title = to_docx['Job Title'].iloc[0]
        annual_spend = to_docx['Department Spend'].iloc[0]
        industry = to_docx['Industry Sector'].iloc[0]
        key_products = to_docx['Key Products or Services'].iloc[0]
        employees = to_docx['Employee Count'].iloc[0]
        revenue = to_docx['Annual Sales'].iloc[0]
        locations = to_docx['Locations '].iloc[0]
        it_count = to_docx['IT Department Size'].iloc[0]
        security_count = to_docx['IT Security Team Size'].iloc[0]
        contact_center = to_docx['Contact Center Seats'].iloc[0]
        op_s = to_docx['Operating System'].iloc[0]
        erp_v = to_docx['Current ERP'].iloc[0]
        cloud_sp = to_docx['Cloud Service Provider'].iloc[0]

        context = {'company': compani,
        'state': state, 
        'annual_spend': annual_spend, 
        'job_title': job_title, 
        'industry': industry,
        'key_products': key_products, 
        'employees': employees, 
        'revenue': revenue, 
        'locations': locations, 
        'it_count': it_count, 
        'security_count': security_count,
        'contact_center': contact_center, 
        'op_s': op_s, 
        'erp_v': erp_v, 
        'cloud_sp': cloud_sp}
        

        # import the word template
        path = '/Users/brych/Documents/Chasetek/Mar22_Show/Chasetek/Mar22_Show/Template.docx'
        doc = DocxTemplate(path)

        # link the variables
        doc.render(context)
        doc.save(f'{company}_report.docx')
    
        return doc

    # figure out if we want the user to be able to select the companies individually or just from the selection
    # add a yes or no line for multiple or only a single company
    # add a multiple choice between the categories for ucaas and all... 
    company_bull = st.radio('Do you want to transfer the current selection to Word doc or just one company?', ('Current Selection', '1 Company')) #current selection

    def download_button(object_to_download, download_filename, button_text, pickle_it=False):
        """
        Generates a link to download the given object_to_download.

        Params:
        ------
        object_to_download:  The object to be downloaded.
        download_filename (str): filename and extension of file. e.g. mydata.csv,
        some_txt_output.txt download_link_text (str): Text to display for download
        link.
        button_text (str): Text to display on download button (e.g. 'click here to download file')
        pickle_it (bool): If True, pickle file.

        Returns:
        -------
        (str): the anchor tag to download object_to_download

        Examples:
     --------
        download_link(your_df, 'YOUR_DF.csv', 'Click to download data!')
        download_link(your_str, 'YOUR_STRING.txt', 'Click to download text!')

        """
        if pickle_it:
            try:
                object_to_download = pickle.dumps(object_to_download)
            except pickle.PicklingError as e:
                st.write(e)
                return None

        else:
            if isinstance(object_to_download, bytes):
                pass

            elif isinstance(object_to_download, pd.DataFrame):
                object_to_download = object_to_download.to_csv(index=False)

            # Try JSON encode for everything else
            else:
                object_to_download = json.dumps(object_to_download)

        try:
            # some strings <-> bytes conversions necessary here
            b64 = base64.b64encode(object_to_download.encode()).decode()

        except AttributeError as e:
            b64 = base64.b64encode(object_to_download).decode()

        button_uuid = str(uuid.uuid4()).replace('-', '')
        button_id = re.sub('\d+', '', button_uuid)

        custom_css = f""" 
            <style>
                #{button_id} {{
                    background-color: rgb(255, 255, 255);
                    color: rgb(38, 39, 48);
                    padding: 0.25em 0.38em;
                    position: relative;
                    text-decoration: none;
                    border-radius: 4px;
                    border-width: 1px;
                    border-style: solid;
                    border-color: rgb(230, 234, 241);
                    border-image: initial;

                }} 
                #{button_id}:hover {{
                    border-color: rgb(246, 51, 102);
                    color: rgb(246, 51, 102);
                }}
                #{button_id}:active {{
                    box-shadow: none;
                    background-color: rgb(246, 51, 102);
                    color: white;
                    }}
            </style> """

        dl_link = custom_css + f'<a download="{download_filename}" id="{button_id}" href="data:file/txt;base64,{b64}">{button_text}</a><br></br>'

        return dl_link

    #if company_bull == 'Current Selection':
       # companies = df_selection['Company'].to_list()
        #button_pdf1 = st.button('Export current selection to Word doc')
        #if button_pdf1: 
           # for c in companies: 
             #   download_button(to_docs_cont(c, selected1_rows))

    if company_bull != 'Current Selection':
        company = st.text_input('Which company do you want to export to Word doc?')
        button_pdf = st.button('Export to Word doc')
        if button_pdf: 
            download_button(to_docs(company, selected_rows), f'{company}_report_c.docx', 'click here to download Word report')

elif authentication_status == False:
    st.error('Username/password is incorrect')

elif authentication_status == None:
    st.warning('Please enter your username and password')
