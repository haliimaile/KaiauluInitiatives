# -*- coding: utf-8 -*-
"""
Created on Tue Jul 21 23:49:57 2026

@author: kyn01

The purpose of this is to create an online interface or dashboard
for Kaiaulu Initiatives. My intention is to create an oluolu, user-
friendly design that is functional and ever green. 
"""


import streamlit as st
import pandas as pd
import plotly.express as px
import pgeocode

import io
from docx import Document              # for downloading a word doc
from docx.enum.section import WD_ORIENT     # for downloading a word doc
from docx.shared import Pt                # for downloading a word doc
from docx.shared import RGBColor

from streamlit_gsheets import GSheetsConnection    # for reading google sheets to streamlit

#import gspread
#from google.oauth2.service_account import Credentials
#from itables.streamlit import interactive_table


# this changes the default font for the whole site, 
st.markdown("""
            <style>
            /* change the whole app's font*/
            /*html, body, [data-testid="stAppViewContainer"],*/
            /*.stApp {*/
             /*   font-family: serif !important;*/
              /*  }*/
                
            
            /* center the main title */
            h1 {
                text-align: center;
            }
            
            /* increase font size*/
            /* targets text from st.write, st.markdown */
           /* [data-testid="stMarkdownContainer"] p { */
           /*     font-size:20px !important;   */
           /*     font-family: serif !important;   */
           /*     }    */
            
            /* increase side bar option/link font-size */
            section[data-testid="stSidebar"] .css-1544g2n,
            section[data-testid="stSidebar"] label,
            section[data-testid="stSidebar"] p {
                font-size: 16px !important;
            }
            </style>
            """, unsafe_allow_html=True
            )


#######################################################################################
# %%


# lists

# rename column names
rsvp_col_names = {
    "When are you joining us for?" : "Date of Workday",
    "What is your first name?" : "First Name",
    "What is your middle name?" : "Middle Name",
    "What is your last name?" : "Last Name",
    "What is your email?" : "Email Address",
    "Would you like to receive email newsletters?" : "Email Newsletter",
    "What is your gender?" : "Gender",
    "How old are you?" : "Age",
    "What is your birthdate?" : "Age",
    "Are you of Native Hawaiian ancestry?" : "Native Hawaiian",
    "What is your street address?" : "Street Address",
    "What state do you live in?" : "State",
    "What city do you live in?" : "City",
    "What is your zip code?" : "Zip Code",
    "Have you signed our waiver since January 1st of this year?" : "Signed Waiver",
    "Are you 18 years or older?" : "Adult",
    "Name of emergency contact" : "Emergency Contact",
    "Phone number of emergency contact" : "Emergency Contact Phone No",
    "Printed name of custodial parent or guardian" : "Name of Guardian",
    "Do you grant us permission to use photographs that may include your image for promotional, educational, or informational purposes?" : "Photo Permission",
    "Initial of custodial parent or guardian" : "Initial of Guardian",
    "Today's date" : "Date of Initials M",         # M = minor
    "Name of emergency contact" : "Emergency Contact",
    "Phone number of emergency contact" : "Emergency Contact Phone No",
    "Do you grant us permission to use photographs that may include your image for promotional, educational, or informational purposes? (2)" : "Photo Permission",
    "Initials for waiver" : "Initials for Waiver",
    "Today's date (2)" : "Date of Initials A",      # A = adult
    "How did you hear about Kaiāulu Initiatives?" : "Heard of KI",
    "Were you or your family impacted by the August 8, 2023 Lāhaina wildfire?" : "Affected by Maui Wildfires"
    
    }

post_col_names = {
    "Today's Date" : "Submission Date",
    "What type of work did you do?" : "Work Done",
    "What did you receive from participating at Kaiāulu? (Select all that apply)" : "Participation Experience",
    "Did you learn anything about the following:" : "Education", 
    "Compared to when you started, how has your overall stress or anxiety level changed?" : "Anxiety Difference",
    #"Which of the following did you experience through Kaiāulu? (Select all that apply)" : "Volunteer Experience",
    "Is there anything else you would like to share about your experience? (We love to hear your accomplishments, too!)" : "Open Ended Share Space"

    }

# columns from df_post to graph 
post_graphable_columns = ["Education", "Anxiety Difference", "Participation Experience"]


# create a dictionary that matches each data field to its best chart type
chart_map = {
    "Email Newsletter" : "Pie",
#    "Email Newsletters?" : "Scatter",
    "Native Hawaiian" : "Pie",
    "Gender": "Pie",
    "Age" : "Bar",
#    "Native Hawaiian?" : "Line",
    "Heard of KI" : "Bar",
    "Affected by Maui Wildfires" : "Bar",
    "Work Done" : "Bar",
    "Participation Experience" : "Bar",
    "Anxiety Difference" : "Bar",
    #"Volunteer Experience": "Bar",
    "Zip Code" : "Bubble Map",
    "Education" : "Bar"
    
    }



# this is column fields to be included in a report, it does not...
# ... include information for the waiver, e.g. date of initials
report_cols = ["Date of Workday", 
               #"Submitted at",
                          "First Name",
                          "Middle Name",
                          "Last Name", 
                          "Email Newsletter", 
                          "Gender", 
                          "Age",
                          "Native Hawaiian", 
                          "Street Address", 
                          "State",
                          "City",
                          "Zip Code", 
                          "Heard of KI", 
                          "Photo Permission",
                          "Affected by Maui Wildfires",
                     
                     ]

# this is a list of columns that should not be graphed and displayed only
display_only_rsvp_cols = ["Date of Workday", 
                          "First Name", 
                          "Middle Name",
                          "Last Name", 
                          "Email Address", 
                          "Street Address",
                          "State", 
                          "City",
                         # "Zip Code", 
                          "Emergency Contact", 
                          "Emergency Contact Phone No", 
                          "Name of Guardian", 
                          "Suffix",
                          "Submission ID", 
                          "Respondent ID",
                          #"Submitted at", 
                          "Signed Waiver",
                          "Photo Permission", 
                          "Minor?",
                          "Initial of Guardian", 
                          "Initials for Waiver",
                          "Date of Initials A", 
                          "Date of Initials M",
                          "Name of emergency contact (2)",
                          "Phone number of emergency contact (2)",
                          "Checked In?",
                          "Open Ended Share Space"  
                          ]


# list all the multiselect question options to explode in a below def
multiselect_columns = ["Received from KI", "Work Done", "Volunteer Experience"]


# link specific questions to specific colors for to aid pattern recognition
color_map = {
    # response to what type of work did you do
    "Watered plants" : "#0b5b26",    # darkish green
    "Planted new plants" : "#3293aa",    # grey-ish blue
    "Weeding" : "#aa5be0",              # violet
    "Reinforced water wells" : "#e7731d",    # rusty orange
    "Removed/cleared invasives" : "#f6c23e",    # lilikoi yellow
    "Water system maintenance" : "#c8192b",     # lipstick red
    "Other" : "#23c7d4",      # cyan
    
    # response choices from "what did you receive from participating w/ KI"
    "Connection to ʻāina/land" : "#94d5e3",     # pastel blue
    "Community and friendship" : "#70c286",     # pastel green
    "Stress relief and peace" : "#e79ac6",       # pastel dark, grey, pink
    "A sense of purpose" : "#c6a4da",            # lavendar
    "Cultural learning" : "#ebcc97",           # pastel tan
    "To give back and help" : "#e99299",       # pastel red
    
    "Yes" : "#66af64",     # green, like KI colors?
    "No" : "#e2770b",      # burnt orange 
    
    # responses from gender question
    "Female" : "#d57bff",        # bright, light purple
    "Male" : "#ffe348",        # yellow, little white washed
    "Non-binary" : "#1aa9c9",      # greyish light blue
    "Decline to state"  : "#d938af",    # violet red
    
    # response to how did you hear of KI?
    "Friend or family" : "#e6edfe",     # grey
    "Social media" : "#fbbf8a",      # light, pastel orange
    "Case manager or service provider" : "#f08d9c",  # dusty rose
    "Community flyer or event" : "#a1d667",    # light pastel green
    
    # response to how has your anxiety/stress changed
    "Much better" : "#309cb4",    # greyish blue
    "Somewhat better" : "#aad2dc",
    "About the same" : "#afdcff",    # light blue
    "Somewhat worse" : "#e2b8ed",      # pastel purple
    "Much worse" : "#a19df9",      # violet
    
    # response to did you learn anything about the following
    "Native plant propagation" : "#f58588",          # dusk pink
    "Native plant care" : "#f8810a",       # orange
    "Fire mitigation" : "#09527c",       # navy blue 
    "History of Lāhaina" : "#78b522",     # lime green
    "Carpentry" : "#de74d8",         # purple
    # other (color already specified above)
    
    
    # response to were you or your family impacted by the August 8, 2023 Lāhaina wildfire
    "Yes, my family or I was directly affected" : "#f2b44f",         # orange-yellow
    "Yes, my family or I were somewhat affected" : "#46a160",          # green
    "My community was impacted" : "#acc8ff",       # dusk blue
    "No direct impact" : "#974fcb",            # purple

    
    }


# %%

# create connection from google sheet to streamlit

@st.cache_data(ttl=60)
def load_gsheet(tab_name):
    conn = st.connection("gsheets", type=GSheetsConnection)
    df = conn.read(worksheet=tab_name, ttl=0)
    df = df.dropna(how="all")      # drop completely empty rows
    return df


# %%

# ------- READ IN RSVP -----------------
df_rsvp = load_gsheet("RSVP")
df_rsvp.rename(columns=rsvp_col_names, inplace=True)


# ------ READ IN POST SURVEY -------------
df_post = load_gsheet("Post-Survey")
df_post.rename(columns=post_col_names, inplace=True)



# strip the olelo from all response choices
# .str.split splits it at the | character
# .str[0] selects which index or which half you want to keep, in this case english
# the str.strip() strips the whitespace which precedes |
# [ WHAT_TO_DO_WITH_x for x in SOMETHING]

for col in df_rsvp.columns:
    df_rsvp[col] = [
        str(response).split("|")[0].strip() if pd.notna(response) else response
        for response in df_rsvp[col]
        ]

# do it again for df_post
for col in df_post.columns:
    df_post[col] = [
        str(response).split("|")[0].strip() if pd.notna(response) else response
        for response in df_post[col]
        ]
    
df_post["Submitted at"] = pd.to_datetime(df_post["Submitted at"], errors="coerce")
    
# clean df_rsvp zip codes
df_rsvp["Zip Code"] = (
    df_rsvp["Zip Code"]
    .str.strip()                    # remove stray spaces
    .str.extract(r"^(\d{5})")[0]    # keep the first 5 digits, or blank if there aren't 5
)
    
    

# now that the olelo has been stripped, the Date of workday is now a string ...
# ... convert it back to datetime
df_rsvp["Date of Workday"] = pd.to_datetime(df_rsvp["Date of Workday"], format = "%Y-%m-%d")


# delete later
#df_rsvp["Submitted at"] = pd.to_datetime(df_rsvp["Submitted at"], format = "%Y-%m-%d %H:%M:%S")


# narrow down df_rsvp to the desired columns, ready to be used in a report tab
df_report = df_rsvp[report_cols]



# %%


# %%

# create defs

 
    
    # col 8 = email , 21/26 = permission to photograph
def show_volunteerlist():
    st.title("Volunteer List")
    
    # value= None means user only chooses one date; table doesn't return more than this date
    chosen_date = st.date_input("Choose date to narrow down options", value=None, 
                  format="MM/DD/YYYY")
                              

    if chosen_date is None:
        st.info("Pick a date to see who signed up.")
        return

    # delete Submitted at, add "Last Name" ####################################################333
    df_voln_list = df_rsvp[["Date of Workday", "First Name", "Last Name",
                            "Photo Permission", "Adult"]]
    
    
    df_voln_list = df_voln_list[df_voln_list["Date of Workday"].dt.date == chosen_date]
    
    # delete later
    #df_voln_list = df_voln_list[df_voln_list["Submitted at"].dt.date == chosen_date] 

    if len(df_voln_list) == 0:
        st.warning("No RSVPs for that date.")
        return

    # checkbox column 
    df_voln_list.insert(0, "Checked In", False)

    # color the row red if photo permission is not "Yes"
    def highlight_no_photo(row):
        if row["Photo Permission"] != "Yes":
            return ["background-color: #f8c9c9"] * len(row)
        return [""] * len(row)

    styled_df = df_voln_list.style.apply(highlight_no_photo, axis=1)

    edited_df = st.data_editor(
        styled_df,
        column_config={
            "Checked In": st.column_config.CheckboxColumn("Checked In", default=False)
        },
        disabled=[col for col in df_voln_list.columns if col != "Checked In"],
        use_container_width=True,
    )
   
# for questions that are multiple select (e.g. work done) that can have ...
# multiple answers separated by comma, use this def to expand the options ...
# ... into mini lists that are not comma separated. 
def explode_if_multiselect(df_report, col_name, multiselect_columns):

    # check to see if it's multicolumn or not
    if col_name not in multiselect_columns:
        return df_report

    split_df = df_report.copy()

    # Split each cell at the commas
    split_df[col_name] = split_df[col_name].str.split(",")

    # Remove the extra spaces around each answer
    cleaned_lists = []
    for answer_list in split_df[col_name]:
        if isinstance(answer_list, list):
            cleaned = []
            for answer in answer_list:
                cleaned.append(answer.strip())
            cleaned_lists.append(cleaned)
        else:
            # empty cells aren't lists, so leave them alone
            cleaned_lists.append(answer_list)
    split_df[col_name] = cleaned_lists

    # Give each answer its own row
    split_df = split_df.explode(col_name)

    return split_df
    

# stream data from google sheets to streamlit
# needs to be called in def build_word_report, def show_report, def volunteer_list
# refresh every 5 min * 60 sec = 300 sec
@st.cache_data(ttl=300)
def stream_data():
    conn = st.connection("gsheets", type=GSheetsConnection)

    # maybe adjust the ttl later on... need to see how it is actually used
    conn_df = conn.read(ttl=60)
    
    # drops rows that are ALL empty
    conn_df.dropna(how="all")
    
    
    return conn_df
                  


# chose word document over pdf because it required less lines of code ...
# ... is editable if need be & hawaiian characters are naturally handled ...
# ... but word docs shift from app to app
def build_word_report(summary_df, report_df, start_text, end_text):
    
    # now make the document
    doc = Document()
    doc.styles["Normal"].font.size = Pt(9)    # adjust font-size
                                                # smaller text so tables can fit better
    
    # turn page sideways so the tables can fit better
    section = doc.sections[0]
    # this only changes a label on the page, needs the subsequent line
    section.orientation = WD_ORIENT.LANDSCAPE
    # tell streamlit the dimensions, turns page (8.5,11) into (11x8.5) -- pair
    section.page_width, section.page_height = section.page_height, section.page_width
        
    # title    
    doc.styles["Heading 1"].font.color.rgb = RGBColor(43,43,41)
    doc.styles["Heading 2"].font.color.rgb = RGBColor(43,43,41)
    
    doc.add_heading("Kaiāulu Initiatives Volunteer Report", level = 1)
    doc.add_heading(f"From {start_text} to {end_text}", level = 2)
    
    doc.add_paragraph()  
    doc.add_paragraph()  
    
    
    # summary statistics table
    # cols = 2 bc only have 2 statistics, can change to len() if we define it later on
    summary_table = doc.add_table(rows=1, cols=2, style = "Light Grid Accent 1")
    summary_table.rows[0].cells[0].text = "Item"
    summary_table.rows[0].cells[1].text = "Value"


    for item, value in zip(summary_df["Item"], summary_df["Value"]):
        # adds a new, empty row at the bottom of the table and hands the row back
        # .cells gets the boxes in that row, one for each column 
        cells = summary_table.add_row().cells
        if isinstance(value, pd.Timestamp):
            value = value.strftime("%B %#d, %Y")
        cells[0].text = str(item)
        cells[1].text = str(value)
        
    doc.add_paragraph()         # add blank line between the 2 tables
    
    # report table
    table_df = report_df.fillna("").astype(str)
    
    report_table = doc.add_table(rows = 1, cols=len(table_df.columns), 
                                 style = "Light Grid Accent 1")
    
    for i, col in enumerate(table_df.columns):
        report_table.rows[0].cells[i].text = col
    for row in table_df.values.tolist():
        cells = report_table.add_row().cells
        for i, value in enumerate(row):
            cells[i].text = value
    
    buffer = io.BytesIO()
    doc.save(buffer)
    return buffer.getvalue()    # the finished word file, as bytes
    
        
@st.fragment
def create_table(df_report_filter):
    st.subheader("Create a report")


    # adjust the column name
    # calculate the amount of volunteers hours worked per time range
    voln_hours = 0
    for date in df_report_filter["Date of Workday"]:
        if date.dayofweek == 2:
            voln_hours += 2.5
        elif date.dayofweek == 6:
            voln_hours += 4
            
    # calculate the number of volunteers per time range
    amnt_voln = len(df_report_filter)
            
    # make the clickable buttons
    # column = col, table = t --> colt
    colt1, colt2, colt3 = st.columns(3)
    
    checkbox_options = list(df_report.columns)
    
    total = len(checkbox_options)
    group_size = total // 3
    leftover_cols = total % 3    # accounts for remainders
    
    first_row = group_size
    second_row = group_size

    
    # make sure that leftover columns go left to right
    if leftover_cols >= 1:
        first_row = first_row + 1
        
    if leftover_cols >= 2:
        second_row = second_row + 1
        
    first_end = first_row
    second_end = first_row + second_row
        
    first_group = checkbox_options[0 : first_end]
    second_group = checkbox_options[first_end : second_end]
    third_group = checkbox_options[second_end :]
    
    
    
    #this list starts empty and fills up with whatever the user checks
    selected_columns = []
    
    # tell streamlit what goes in column 1 
    with colt1:
        for column in first_group:
            if st.checkbox(column, key = f"report_{column}"):
                selected_columns.append(column)
                
    with colt2:
        for column in second_group:
            if st.checkbox(column, key = f"report_{column}"):
                selected_columns.append(column)
                  
    with colt3:
        for column in third_group:
            if st.checkbox(column, key = f"report_{column}"):
                selected_columns.append(column)
                
    if not selected_columns:
        st.info("Check at least one box to build a report table.")
        return
    
    
    st.space(size="small")
    
    stats1_col, stats2_col = st.columns(2)
    
    st.space(size = "small")
    
    stats1_col.metric("Volunteers", f"{amnt_voln}", border = True)
    stats2_col.metric("Volunteer Hours", f"{voln_hours}",border = True)
    
    
    
    if selected_columns:    # if they exist then...
        
        printable_df = df_report_filter[selected_columns].copy()
        
        for col in printable_df.select_dtypes(include=["datetime"]).columns:
            printable_df[col] = printable_df[col].dt.strftime("%B %#d, %Y")
            
        display_df = df_report_filter[selected_columns]
        st.dataframe(
            display_df,
            use_container_width=True,
            column_config={
                "Date of Workday": st.column_config.DateColumn(format="MMMM D, YYYY")
                },
            )
        
        
        
        #st.dataframe(printable_df, use_container_width=True)
        
        if len(df_report_filter) > 0:
            start_text = df_report_filter['Date of Workday'].min().strftime("%B %#d, %Y")
            end_text = df_report_filter['Date of Workday'].max().strftime("%B %#d, %Y")
   
        # delete later
        #if len(df_report_filter) > 0:
        #    start_text = df_report_filter['Submitted at'].min().strftime("%B %#d, %Y")
        #    end_text = df_report_filter['Submitted at'].max().strftime("%B %#d, %Y")
        #    df_rsvp["Submitted at"] = pd.to_datetime(df_rsvp["Submitted at"], format = "%Y-%m-%d %H:%M:%S")
            
            # start building the report here
            # make the small table summarizing statistics
            summ_df = pd.DataFrame({
                "Item" : ["Number of Volunteers", "Volunteer hours"],
                "Value" : [amnt_voln, voln_hours]
                })

            # turn summary table into csv text
            summary_csv = summ_df.to_csv(index=False)
            table_csv = printable_df.to_csv(index=False)
            
            # turn the table into csv text that the button can hand to the user
            csv_data = (summary_csv + "\n" + table_csv).encode("utf-8")
    
            # turn table into word doc
            word_bytes = build_word_report(summ_df, printable_df, start_text, end_text)

            word_col, csv_col = st.columns(2)
        
            with word_col:
                st.download_button(
                    label="Download WORD",
                    data = word_bytes,
                    file_name=f"report_{start_text}to{end_text}.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    key = "download_word",
                    width = "stretch"
                        )
            
            with csv_col:
                st.download_button(
                    label="Download CSV",
                    data = csv_data,
                    file_name=f"report_{start_text}to{end_text}.csv",
                    mime="text/csv",
                    width = "stretch"   # makes the button fill the column
                    )

        else:
            st.info("Check at least one box to build a report.")
            
def show_post_results(df_post, start_date, end_date):
    st.header("Post-Event Survey Results")

    start = pd.to_datetime(start_date)
    end = pd.to_datetime(end_date) + pd.Timedelta(days=1)

    df_post_filter = df_post[
        (df_post["Submitted at"] >= start) &
        (df_post["Submitted at"] < end)
    ]

    if len(df_post_filter) == 0:
        st.info("No survey responses were submitted in this date range.")
        return
    
    post_chart_map = {
    "Education": "Bar",
    "Anxiety Difference": "Pie",
    "Participation Experience": "Bar",
    "Work Done" : "Bar"
}

    post_graphable_columns = ["Education", "Anxiety Difference", 
                              "Participation Experience", "Work Done"]

    for column in post_chart_map:
        chart_type = post_chart_map[column]
        
        st.subheader(column, text_alignment="center")

        working_df = explode_if_multiselect(df_post_filter, column, multiselect_columns)
        counts = working_df[column].value_counts().reset_index()
        counts.columns = [column, "Number"]

        if chart_type == "Pie":
            fig = px.pie(
                counts,
                names=column,
                values="Number",
                color=column,
                color_discrete_map=color_map
            )
            fig.update_traces(
                textfont=dict(size=16, color="black"),
                textinfo="label+value"
            )

        if chart_type == "Bar":
            fig = px.bar(
                counts,
                x=column,
                y="Number",
                color=column,
                color_discrete_map=color_map,
                text="Number"
            )
            fig.update_traces(textfont=dict(size=16, color="black"))
            fig.update_layout(showlegend=False)
        
        st.plotly_chart(fig, use_container_width=True)

    # open-ended answers, shown as a list instead of a chart
    st.subheader("Shared Responses")
    shared = df_post_filter[["Open Ended Share Space"]].dropna()
    st.dataframe(shared, use_container_width=True)


# kind of misleading title... this shows the charts
def show_reports(df_report):
    st.title("Reports")
    
    # start by creating a date range for the user to select from
    st.subheader("Select a date range")
    date_range = st.date_input("Select a start and end date",
        value=[], format="MM/DD/YYYY")

    if len(date_range) != 2:
        st.info("Please select a start and end date, even if it's the same date.")
        return

    start_date, end_date = date_range
   
    
   # use this when dates are recorded
    df_report_filter = df_report[
       (df_report["Date of Workday"] >= pd.to_datetime(start_date)) &
       (df_report["Date of Workday"] <= pd.to_datetime(end_date))
       ]
    
    
    # temporary use -- delete later
    #df_report_filter = df_report[
    #   (df_report["Submitted at"] >= pd.to_datetime(start_date)) &
    #   (df_report["Submitted at"] <= pd.to_datetime(end_date))
    #   ]
    
    
    # report table   
    st.divider()   ############################################################
    
    
    create_table(df_report_filter)
    
   
    # graphing area
    st.divider()   ############################################################


    st.subheader("Select what to visualize")

    graphable_columns = [col for col in df_report.columns if col not in display_only_rsvp_cols]
    total_cols = len(graphable_columns)   # total amount
    
    #this list starts empty and fills up with whatever the user checks
    selected_columns_multi = []
    
    # make 3x2 table of checkboxes    
    # define the name of each column to call upon it specifically
    col1, col2, col3 = st.columns(3)
    
    group_size = total_cols // 3
    leftover_cols = total_cols % 3    # accounts for remainders
    
    first_size = group_size
    second_size = group_size

    # make sure that leftover columns go left to right
    if leftover_cols >= 1:
        first_size = first_size + 1
        
    if leftover_cols >= 2:
        second_size = second_size + 1
        
    first_end = first_size
    second_end = first_size + second_size

    
    # specify the rows 
    row_one = graphable_columns[0 : first_end]
    row_two = graphable_columns[first_end : second_end]
    row_three = graphable_columns[second_end :]
    

    
    # tell streamlit what goes in column 1 
    with col1:
        for column in row_one:        # [0:len(graphable_columns) // 3]     12 graphable_cols // 3 = 4 --> 0:4
            if st.checkbox(column, key = f"chart_{column}"):
                selected_columns_multi.append(column)
                
    with col2:
        for column in row_two:      # [(len(graphable_columns) // 3) : (len(graphable_columns) // 3) +]
            if st.checkbox(column, key = f"chart_{column}"):
                selected_columns_multi.append(column)
                  
    with col3:
        for column in row_three:
            if st.checkbox(column, key = f"chart_{column}"):
                selected_columns_multi.append(column)
                
    if not selected_columns_multi:
        st.info("Check at least one box to see a chart.")
    else:
        for column in selected_columns_multi:

            # displays the desired chart
            chart_type = chart_map.get(column)
        
            # lets the user know what the computer is doing
            st.subheader(f"{column}", text_alignment="center")
    
        
    
            if chart_type == "Pie":
                working_df = explode_if_multiselect(df_report_filter, column, multiselect_columns)
                counts = working_df[column].value_counts().reset_index()
                counts.columns = [column, "Number"]
                fig = px.pie(
                    counts, 
                    names=column, 
                    values="Number",
                    color=column,
                    color_discrete_map=color_map
                    )
            
                fig.update_traces(
                    textfont=dict(size=16, color = "black"),
                    textinfo = "label+value"     # shows the count instead of default %
                    )

        
            elif chart_type == "Bar":
                working_df = explode_if_multiselect(df_report_filter, column, multiselect_columns)
                #counts = working_df[column].value_counts().reset_index()
                #counts.columns = [column, "Count"]
            
                if column == "Age":
                    # delete when you get the data that explicitly lists ages
                    #birth_dates = pd.to_datetime(working_df["Age"], errors="coerce")
                    #ages = (working_df["Date of Workday"] - birth_dates).dt.days // 365.25
                    ages = working_df['Age']
                
                
                    age_bins = [0, 13, 18, 25, 35, 50, 65, 200]
                    age_labels = ["12 and under", "13-17", "18-24",
                                  "25-34", "35-49", "50-64", "65+"]
                
                    ages = pd.to_numeric(ages, errors="coerce")

                    working_df[column] = pd.cut(ages, bins=age_bins,
                            labels=age_labels, right=False)
                
                    working_df = working_df.copy()
                    # pd.cut allows numerical data to be divided into ranges
                    working_df[column] = pd.cut(ages, bins=age_bins,
                                                labels= age_labels, right=False)
                
                    counts = working_df[column].value_counts(sort=False).reset_index()
                
                else:
                    # every other chart does exactly what it did before
                    counts = working_df[column].value_counts().reset_index()

                counts.columns = [column, "Number"]
            
                fig = px.bar(
                    counts, 
                    x=column, 
                    y="Number",
                    color=column,
                    color_discrete_map=color_map,
                    text = "Number"
                    )
            
                # format tick step size based off total # of data
                max_count = counts["Number"].max()
                step = max(1, round(max_count / 10))    # desiring about 10 ticks
            
                # format the graph layout
                fig.update_layout(
                    legend_title_font_color="#333333",
                    font=dict( 
                    color = "#333333",
                    size = 16
                        ),
                    xaxis = dict(
                        tickfont=dict(size=14, color = "#333333"),
                        title = dict(font=dict(size = 16, color = "#333333"))
                        ),
                    yaxis = dict(
                        dtick = step,
                        tickformat="d",
                        color = "#333333",
                        tickfont=dict(size=14, color = "#333333"),
                        title = dict(font=dict(size = 16, color = "#333333"))
                        ),
                    legend = dict(
                        font= dict(size=14),
                        title=dict(font=dict(size=18, color = "#333333"))
                        )
                    )
            # labels on the bar
                fig.update_traces(textfont=dict(size=14, color = "white"))
    
            elif chart_type == "Line":
                fig = px.line(df_report_filter, y=column)
                
                
            elif chart_type == "Bubble Map":
                # count how many volunteers are in each zip code
                counts = df_report_filter[column].value_counts().reset_index()
                counts.columns = [column, "Number"]
                
                # start setting up: looking up geographical info from zip code
                zip_lookup = pgeocode.Nominatim("US")
                
                # look up every zip code in zip_lookup at once
                places = zip_lookup.query_postal_code(counts[column].tolist())
                
                # add coordinates to the counts, row for row
                counts['Latitude'] = places['latitude'].values
                counts['Longitude'] = places['longitude'].values
                
                # separate the zip codes we could locate vs the ones we couldn't
                placed = counts.dropna(subset=["Latitude", "Longitude"])
                unplaced = counts[counts["Latitude"].isna()]
                
                # let the user know which zip codes could not be located
                if len(unplaced) > 0:
                    st.caption(f"{unplaced['Number'].sum()} volunteer(s) are not shown "
                           "because their zip code could not be located.")
                
                # back up notification, let user know if none of the zip codes could be placed
                if placed.empty:
                    st.info("None of the zip codes could be placed on the map.")
                    continue
                
                # draw bubble map
                fig = px.scatter_map(
                    placed,
                    lat = "Latitude",
                    lon = "Longitude",
                    size = "Number",
                    size_max = 40,
                    hover_name = column,
                    hover_data={"Number": True, "Latitude": False, "Longitude": False},
                    center=dict(lat=20 + 54/60 + 22.13/3600,
                                lon = -(156 + 41/60 + 03.25/3600)),
                    zoom = 9,
                    map_style = "open-street-map",
                    height = 600    # default is ~450
                    )
                
                
        
            st.plotly_chart(fig, use_container_width=True, 
                            key = f"{column}_{chart_type}")
            
            
def show_links():
    # Title of the web page
    st.title("Quick Links")
    
    st.space()
    st.space() 
    
    col1, col2 = st.columns(2)
     
    with col1:
        st.link_button("Enter RSVP", url = "https://tally.so/r/440jrA")
            

    with col2:
        st.link_button("Enter Post-Survey", url = "https://tally.so/r/vG62Yl")

        


# %%



st.set_page_config(
    page_title="Kaiāulu Initiatives Volunteer Dashboard",
    #page_icon="🌳",
    page_icon="https://welocreations.com/cdn/shop/files/IMG_1242_2048x.jpg?v=1690349982",
    layout="wide",
)


 
# create a sidebar menu
# 'menu' is name of the sidebar, but can be changed to anything
st.sidebar.title('Menu')

# create pages within the sidebar menu
# anything to do with sidebar needs to include the .sidebar
# the pages/options within a sidebar menu is referred to as 'radio' here
# 'pages' is what we label our sidebar options. the options are a list of their names
options = st.sidebar.radio("Navigation", options=["Volunteer List", "Reports", "Links"])


# what to show on the pages
if options == "Volunteer List":
    show_volunteerlist()
elif options == "Reports":
    show_reports(df_report)
elif options == "Links":
    show_links()
    



























