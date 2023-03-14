from __future__ import print_function

import os.path
import pickle
import pandas as pd
import numpy as np

import gspread
from gspread_dataframe import set_with_dataframe
from pydrive.auth import GoogleAuth
from pydrive.drive import GoogleDrive

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

SCOPES = ['https://www.googleapis.com/auth/spreadsheets']
SPREADSHEET_ID = '1TnwuYCViBkKxFFkockuDx7Dp9fe-GZp0cfvEQs3mFD4'
APPLICANT_RANGE = 'Input (Raw Data)'
QUESTION_RANGE = 'Questions!A:B'
INDICATOR_RANGE = "Indicators!"
GRADE_RANGE = "Grades"


# User must update this variable with path to json file
path = 'path/to/file.json'

# Update the cell numbers here from the Indicators Tab
gender_cells = "A19:B24"
ethnicity_race_cells = "G2:H12" 
parent_education_cells = "G16:H27" 
gpa_cells = "D27:E33" 
math_science_grades_cells = "J2:N3" 
school_district_cells = "D19:E21" 
lgbtq_identifying_cells = "J8:K11"
trans_identifying_cells =  "M8:N11"
income_cells = "G31:H34" 
reviewer_cells = "J15:J23" 

question_cells = "A2:B40"

c_exp_cells = "B2:E2"
java_exp_cells = "B3:E3"
js_exp_cells = "B4:E4"
py_exp_cells = "B5:E5"
html_exp_cells = "B6:E6"
cyber_exp_cells = "B7:E7"
robot_exp_cells = "B8:E8"
gamedev_exp_cells = "B9:E9"
ai_exp_cells = "B10:E10"
hardware_exp_cells = "B11:E11"
vr_exp_cells = "B12:E12"
scratch_exp_cells = "B13:E13"
google_exp_cells = "B14:E14"

science_grade_cells ="K2:N2"
math_grade_cells = "K3:N3"

default_gender_score = 1
default_ethnicity_race_score = 1
default_parent_education_score = 0
default_gpa_score = 0
default_school_district_score = 0
default_income_score = 0
default_lgbtq_identifying_score = 0
default_trans_identifying_score = 0
default_experience_score = 0

def main():
    """Shows basic usage of the Sheets API.
    Prints values from a sample spreadsheet.
    """
    creds = None
    # The file token.json stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.
    if os.path.exists('token.json'):
        creds = Credentials.from_authorized_user_file('token.json', SCOPES)
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open('token.json', 'w') as token:
            token.write(creds.to_json())

    try:
        service = build('sheets', 'v4', credentials=creds)

        # Call the Sheets API
        sheet = service.spreadsheets()
        
        process_data(sheet)

    except HttpError as err:
        print(err)

def import_from_sheet(sheet, data_range, type):
    values = sheet.values().get(spreadsheetId=SPREADSHEET_ID,
            range=data_range).execute().get('values')
    if not values:
        print("No data found")
    else:
        if type == "dict":
            return dict(values)
        elif type == "list":
            return list(values)
        else:
            print("Have to provide list or dict as argument")
            return

def export_to_sheet(df):
    scopes = ['https://www.googleapis.com/auth/spreadsheets',
            'https://www.googleapis.com/auth/drive']

    credentials = Credentials.from_service_account_file(path, scopes=scopes)

    gc = gspread.authorize(credentials)

    gauth = GoogleAuth()
    drive = GoogleDrive(gauth)

    # open a google sheet
    gs = gc.open_by_key(SPREADSHEET_ID)
    # select a work sheet from its name
    worksheet1 = gs.worksheet('Output')

    # write to dataframe
    # worksheet1.clear()
    set_with_dataframe(worksheet=worksheet1, dataframe=df, include_index=False,
    include_column_header=True, resize=True)

def process_data(sheet):
    applicant_data = import_from_sheet(sheet, APPLICANT_RANGE, "list")
    question_data = import_from_sheet(sheet, QUESTION_RANGE, "dict")
    grade_data = import_from_sheet(sheet, GRADE_RANGE, "list")
    
    parent_education_dict = import_from_sheet(sheet, INDICATOR_RANGE+parent_education_cells, "dict")
    income_dict = import_from_sheet(sheet, INDICATOR_RANGE+income_cells, "dict")
    school_district_dict = import_from_sheet(sheet, INDICATOR_RANGE+school_district_cells, "dict")
    ethnicity_race_dict = import_from_sheet(sheet, INDICATOR_RANGE+ethnicity_race_cells, "dict")
    gender_dict = import_from_sheet(sheet, INDICATOR_RANGE+gender_cells, "dict")
    lgbtq_identifying_dict = import_from_sheet(sheet, INDICATOR_RANGE+lgbtq_identifying_cells, "dict")
    trans_identifying_dict = import_from_sheet(sheet, INDICATOR_RANGE+trans_identifying_cells, "dict")
    gpa_dict = import_from_sheet(sheet, INDICATOR_RANGE+gpa_cells, "dict")

    def exp_dict_generator(exp_array):
        # Hacky fix for now; come back to this later
        exp_array = exp_array[0]
        d = dict()
        d["No Experience"] = float(exp_array[0])
        d["Somewhat Experienced"] = float(exp_array[1])
        d["Experienced"] = float(exp_array[2])
        d["Very Experienced"] = float(exp_array[3])
        return d


    c_dict = exp_dict_generator(import_from_sheet(sheet, INDICATOR_RANGE+c_exp_cells, "list"))
    java_dict = exp_dict_generator(import_from_sheet(sheet, INDICATOR_RANGE+java_exp_cells, "list"))
    js_dict = exp_dict_generator(import_from_sheet(sheet, INDICATOR_RANGE+js_exp_cells, "list"))
    py_dict = exp_dict_generator(import_from_sheet(sheet, INDICATOR_RANGE+py_exp_cells, "list"))
    html_dict = exp_dict_generator(import_from_sheet(sheet, INDICATOR_RANGE+html_exp_cells, "list"))
    cyber_dict = exp_dict_generator(import_from_sheet(sheet, INDICATOR_RANGE+cyber_exp_cells, "list"))
    robot_dict = exp_dict_generator(import_from_sheet(sheet, INDICATOR_RANGE+robot_exp_cells, "list"))
    gamedev_dict = exp_dict_generator(import_from_sheet(sheet, INDICATOR_RANGE+gamedev_exp_cells, "list"))
    ai_dict = exp_dict_generator(import_from_sheet(sheet, INDICATOR_RANGE+ai_exp_cells, "list"))
    hardware_dict = exp_dict_generator(import_from_sheet(sheet, INDICATOR_RANGE+hardware_exp_cells, "list"))
    vr_dict = exp_dict_generator(import_from_sheet(sheet, INDICATOR_RANGE+vr_exp_cells, "list"))
    scratch_dict = exp_dict_generator(import_from_sheet(sheet, INDICATOR_RANGE+scratch_exp_cells, "list"))
    google_dict = exp_dict_generator(import_from_sheet(sheet, INDICATOR_RANGE+google_exp_cells, "list"))


    # Set column titles to be the 2nd row instead of the 0th row because Qualtrics adds a row with "Q1", etc before the row of questions
    df = pd.DataFrame(applicant_data[2:], columns=applicant_data[1])
    pd.set_option("display.max_rows", None)

    question_dict = {v: k for k, v in question_data.items()}

    # Remove title from question list
    title = next(iter(question_dict))
    question_dict.pop(title)

    # print(df.columns)

    # Adding the other option to the selected option here for simplicity. Will fix this with a catch all questiion in the future
    # Potential downside: answers without an other have a blank space afterwards now
    courses_selected = "Select all math and computer science courses that you will have completed by June 30, 2023. - Selected Choice"
    courses_other = "Select all math and computer science courses that you will have completed by June 30, 2023. - Other: - Text"
    programs_selected = "Have you participated in any of the following events and/or programs with the UMD, Iribe Initiative, and other organizations?\n\n(Select all that apply) - Selected Choice"
    programs_other = "Have you participated in any of the following events and/or programs with the UMD, Iribe Initiative, and other organizations?\n\n(Select all that apply) - Other: - Text"

    df[courses_selected] = df[courses_selected] + " " + df[courses_other]
    df[programs_selected] = df[programs_selected] + " " + df[programs_other] 


    # Only keep the questions from the google sheets list in the dataframe of all applicant data
    df = df[question_dict.keys()]

    df.rename(columns=question_dict, inplace=True)

    # Starting at 2 bc of extraneous row; change later
    grade_df = pd.DataFrame(grade_data, columns=grade_data[1])
    grade_df.drop([0,1], inplace=True)
    # Rename these once I recreate grade tab
    grade_df = grade_df[["Student First Name:", "Student Last Name:", "Math", "Science"]]
    grade_df.rename(columns={"Student First Name:": "Student First Name", "Student Last Name:": "Student Last Name", "Math": "Math Grade", "Science":"Science Grade"}, inplace=True)


    df = pd.merge(df, grade_df, on=["Student First Name", "Student Last Name"], how="outer")
    df["Math Grade"].replace([np.nan, r'^\s*$'], "NG", regex=True, inplace=True)
    df["Science Grade"].replace([np.nan, r'^\s*$'], "NG", regex=True, inplace=True)
    

    education_title = "Parent's highest level of education"
    df[education_title] = df[education_title].apply(lambda x: parent_education_dict[x] if x in parent_education_dict else default_parent_education_score)

    income_title = "Low-income status"
    df[income_title] = df[income_title].apply(lambda x: income_dict[x] if x in income_dict else default_income_score)

    school_district_title = "School District/County"
    df[school_district_title] = df[school_district_title].apply(lambda x: school_district_dict[x] if x in school_district_dict else default_school_district_score)

    ethnicity_race_title = "Race/Ethnicity"
    df[ethnicity_race_title] = df[ethnicity_race_title].apply(lambda x: ethnicity_race_dict[x] if x in ethnicity_race_dict else default_ethnicity_race_score)

    gpa_title = "GPA"
    df["GPA (1-4)"] = df[gpa_title].apply(lambda x: gpa_dict[x] if x in gpa_dict else default_gpa_score)

    def calculate_gender_score(gender_response, lgbtq_response, trans_response):
        total_score = 0

        if gender_response in gender_dict:
            total_score += int(gender_dict[gender_response])
        else:
            total_score += default_gender_score

        if lgbtq_response in lgbtq_identifying_dict:
            total_score += int(lgbtq_identifying_dict[lgbtq_response])
        else:
            total_score += default_lgbtq_identifying_score

        if trans_response in trans_identifying_dict:
            total_score += int(trans_identifying_dict[trans_response])
        else:
            total_score += default_trans_identifying_score
        return total_score

    gender_title = "Gender"
    lgbtq_title = "LGBTQ-Identifying"
    trans_title = "Trans-Identifying"
    df["Gender (0-3)"] = df.apply(lambda x: calculate_gender_score(x[gender_title], x[lgbtq_title], x[trans_title]), axis=1)


    c_title = "Experience Question: C# or C++"
    df[c_title] = df[c_title].apply(lambda x: c_dict[x] if x in c_dict else default_experience_score)

    java_title = "Experience Question: Java"
    df[java_title] = df[java_title].apply(lambda x: java_dict[x] if x in java_dict else default_experience_score)

    js_title = "Experience Question: JavaScript"
    df[js_title] = df[js_title].apply(lambda x: js_dict[x] if x in js_dict else default_experience_score)
    
    py_title = "Experience Question: Python"
    df[py_title] = df[py_title].apply(lambda x: py_dict[x] if x in py_dict else default_experience_score)

    html_title = "Experience Question: HMTL/CSS"
    df[html_title] = df[html_title].apply(lambda x: html_dict[x] if x in html_dict else default_experience_score)

    cyber_title = "Experience Question: Cybersecurity"
    df[cyber_title] = df[cyber_title].apply(lambda x: cyber_dict[x] if x in cyber_dict else default_experience_score)

    robot_title = "Experience Question: Robotics"
    df[robot_title] = df[robot_title].apply(lambda x: robot_dict[x] if x in robot_dict else default_experience_score)

    gamedev_title = "Experience Question: Game Development"
    df[gamedev_title] = df[gamedev_title].apply(lambda x: gamedev_dict[x] if x in gamedev_dict else default_experience_score)

    ai_title = "Experience Question: Artificial Intelligence/Machine Learning"
    df[ai_title] = df[ai_title].apply(lambda x: ai_dict[x] if x in ai_dict else default_experience_score)

    hardware_title = "Experience Question: Hardware"
    df[hardware_title] = df[hardware_title].apply(lambda x: hardware_dict[x] if x in hardware_dict else default_experience_score)

    vr_title = "Experience Question: Virtual Reality"
    df[vr_title] = df[vr_title].apply(lambda x: vr_dict[x] if x in vr_dict else default_experience_score)

    scratch_title = "Experience Question: Scratch/Block-Based/Drag and Drop"
    df[scratch_title] = df[scratch_title].apply(lambda x: scratch_dict[x] if x in scratch_dict else default_experience_score)

    google_title = "Experience Question: Microsoft/Google Suite"
    df[google_title] = df[google_title].apply(lambda x: google_dict[x] if x in google_dict else default_experience_score)

    exp_column_names = [c_title, java_title, js_title, py_title, html_title, cyber_title, robot_title, gamedev_title, ai_title, hardware_title, vr_title, scratch_title, google_title]
    df["Experience Score (0-100)"] = df[exp_column_names].sum(axis=1)


    columns_to_drop = {
        'GPA',
        'Gender',
        'Trans-Identifying', 
        'LGBTQ-Identifying', 
        'Experience Question: C# or C++',
        'Experience Question: Java', 
        'Experience Question: JavaScript',
        'Experience Question: Python', 
        'Experience Question: HMTL/CSS',
        'Experience Question: Cybersecurity', 
        'Experience Question: Robotics',
        'Experience Question: Game Development',
        'Experience Question: Artificial Intelligence/Machine Learning',
        'Experience Question: Hardware', 
        'Experience Question: Virtual Reality',
        'Experience Question: Scratch/Block-Based/Drag and Drop',
        'Experience Question: Microsoft/Google Suite'
    }

    df.drop(columns_to_drop, axis="columns", inplace = True)

    df["Applicant ID"] = df.reset_index().index

    reviewer_columns = [
        "Reviewer 1",
        "Reviewer 1 Decision",
        "Reviewer 1 Notes",
        "Reviewer 2",
        "Reviewer 2 Decision",
        "Reviewer 2 Notes",
        "Reviewer 3",
        "Reviewer 3 Decision",
        "Reviewer 3 Notes",
        "Final Decision",
    ]

    for col in reviewer_columns:
        df[col] = None

    column_order = [
        "Applicant ID",
        "Reviewer 1",
        "Reviewer 1 Decision",
        "Reviewer 1 Notes",
        "Reviewer 2",
        "Reviewer 2 Decision",
        "Reviewer 2 Notes",
        "Reviewer 3",
        "Reviewer 3 Decision",
        "Reviewer 3 Notes",
        "Final Decision",
        'Student First Name', 
        'Student Last Name', 
        'Create Tech Title', 
        'AI Camp Title', 
        'Cyber Title',
        'Current Grade Level',
        'Gender (0-3)',
        'Race/Ethnicity',
        'GPA (1-4)', 
        'Math Grade',
        'Science Grade', 
        'School District/County', 
        'Parent\'s highest level of education', 
        'Low-income status',
        'Experience Score (0-100)',
        'Math/Technology Courses', 
        'Other Technology Courses',
        'Hobbies & Extracurriculars', 
        'Prior participation in our programs',
        'Use this space to share any links to your work. (optional)',
        'Use this space to share any additional information with us.',
        'Accomodations', 
        'Why you want to attend the program.',
        'Essay-Reponse Question about Technology (Not program specific)',
        'Create Tech-Specific Question', 
        'Cyber Defense-Specific Question 1',
        'Cyber Defense-Specific Question 2',
        'AI Summer Program-Specific Question',
        'Name of 1st Reference', 
        'Name of 2nd Reference',
    ]

    df = df[column_order]

    export_to_sheet(df)

    return

if __name__ == '__main__':
    main()