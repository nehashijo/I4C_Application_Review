import pickle
import os.path
import pandas as pd

from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from googleapiclient.discovery import build

import gspread
from gspread_dataframe import set_with_dataframe
from google.oauth2.service_account import Credentials
from pydrive.auth import GoogleAuth
from pydrive.drive import GoogleDrive



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


def gsheet_api_check(SCOPES):
    creds = None
    if os.path.exists('token.pickle'):
        with open('token.pickle', 'rb') as token:
            creds = pickle.load(token)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        with open('token.pickle', 'wb') as token:
            pickle.dump(creds, token)
    return creds

def pull_sheet_data(SCOPES,SPREADSHEET_ID,DATA_TO_PULL):

    creds = gsheet_api_check(SCOPES)
    service = build('sheets', 'v4', credentials=creds)
    sheet = service.spreadsheets()
    result = sheet.values().get(
        spreadsheetId=SPREADSHEET_ID,
        range=DATA_TO_PULL).execute()
    values = result.get('values', [])
    
    if not values:
        print('No data found.')
    else:
        rows = sheet.values().get(spreadsheetId=SPREADSHEET_ID,
                                  range=DATA_TO_PULL).execute()
        data = rows.get('values')
        print("COMPLETE: Data copied")
        return data

def load_values_as_dict(sheet, cells):

    sheet_range = sheet + "!" + cells

    data = pull_sheet_data(SCOPES,SPREADSHEET_ID,sheet_range)
    dictionary = dict(data)
    return dictionary

def load_values_as_array(cells):

    sheet_range = "Indicators!" + cells

    data = pull_sheet_data(SCOPES,SPREADSHEET_ID,sheet_range)
    return data[0]

SCOPES = ['https://www.googleapis.com/auth/spreadsheets']
SPREADSHEET_ID = '1TnwuYCViBkKxFFkockuDx7Dp9fe-GZp0cfvEQs3mFD4'
DATA_TO_PULL = 'Input (Raw Data)'
data = pull_sheet_data(SCOPES,SPREADSHEET_ID,DATA_TO_PULL)

# Set column titles to be the 1st row instead of the 0th row because Qualtrics adds a row with "Q1", etc before the row of questions
df = pd.DataFrame(data[1:], columns=data[1])

pd.set_option("display.max_rows", None)

# print(df)

rev_question_dict = load_values_as_dict("Questions", question_cells)

question_dict = {v: k for k, v in rev_question_dict.items()}

print(question_dict['Describe your current hands-on computing experience.  Please note experience is not required for most programs. - Virtual Reality/Augmented Reality'])

# df.rename(columns=question_dict, inplace=True)

# print(list(question_dict.values()))


df = df.loc[:, list(question_dict.keys())]

print(df)

# print(df['Experience Question: Virtual Reality'])


# print(list(df.columns.values))



columns_to_drop = [
    "Timestamp", 
    "Student Email Address (Personal Preferred)", 
    "Parent or Guardian First & Last Name", 
    "Parent or Guardian Phone Number", 
    "Parent or Guardian Email Address", 
    "Mailing Address ", 
    "City", 
    "What state do you live in?", 
    "If you selected a \"U.S. Terrorities\" or \"Outside of the U.S. & U.S. Terrorities\", please provide your address below:",
    "Zip Code",
    "Are you fully vaccinated for COVID?",
    "Age by July 1st, 2022?",
    "Student's Preferred Pronouns",
    "School Name (Spring 2022)",
    "School's State",
    "School Type",
    "Are there any learning, behavioral, or physical challenges we should know about to best support your  overall learning experience?",
    "Do you have any dietary restrictions or Allergies?",
    "T-Shirt Size",
    "How did you hear about our summer programs? (Select all that apply)",
    "Title of Teacher Reference ",
    "Email of Teacher Reference:",
    "Title of 2nd Teacher or Counselor Reference",
    "Email of 2nd Teacher or Counselor Reference:",
    "As the parent/ guardian, I certify that my child is a good candidate for attending this summer program, and I will do all I can to support their interest in computing",
    "As the parent/ guardian, I understand my child will be expected to respect his/her peers, guest instructors, and program staff. I will review the behavior contract with my child at the beginning of the summer session.",
    "Signature of Parent/Guardian"
]

columns_to_rename = {
    "Grade Level (Spring 2022)": "Current Grade Level",
    "I am applying to be considered for the following Iribe Initiative Summer Academy high school program(s):": "Programs Applied For",
    "Describe why you want to attend and what you hope to gain from this program.": "Why you want to attend the program.",
    "Describe what interests you about computing.": "What interests you about computing?",
    "Describe how you would solve a problem you care about using technology (this could be a social problem, a technical problem, a local problem, a world problem, etc.).": "Solve a problem using technology",
    "Describe any challenges with technological advances that you see and how it can impact society. ": "Challenges with technological advances in society",
    "Describe your future educational and career goals.": "Future career & educational goals.",
    "List any relevant math you have taken in middle and/or high school.": "Relevant middle/high school math classes",
    "List any relevant cybersecurity, technology, and computing courses you have taken in middle and/or high school.": "Middle/high school technology classes",
    "Please list the highest math class offered at your school.": "Highest math class offered at school",
    "List other computing experiences, hobbies or extra curricular activities you enjoy in addition to information provided above. ": "Hobbies & Extracurriculars",
    "Have you participated in any of our outreach programs or with other organizations? (Select all that apply)": "Prior participation in our programs",
    "Name of Math, Science, or Technology Teacher Reference:": "Name of 1st Reference",
    "Name of  2nd Teacher or Counselor Reference:": "Name of 2nd Reference"
    }
 
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

# def calculate_grade_score(science_response, math_response):
#     total_score = 0
#     if science_response in science_grade_dict and math_response in math_grade_dict:
#         total_score += int(science_grade_dict[science_response])
#         total_score += int(math_grade_dict[math_response])
#         return total_score/2
#     else:
#         return "Grades Missing"

df.drop(columns_to_drop, axis="columns", inplace = True)

df.rename(columns=columns_to_rename, inplace=True)



df["Science Grade"]="Must Manually Input Grade Data"
df["Math Grade"]="Must Manually Input Grade Data"

parent_education_dict = load_values_as_dict("Indicators", parent_education_cells)
education_title = "What is the highest level of education any of your parents has achieved?"
df["Parent's highest level of education (0-2)"] = df[education_title].apply(lambda x: parent_education_dict[x] if x in parent_education_dict else default_parent_education_score)

income_dict = load_values_as_dict("Indicators", income_cells)
income_title = "Do you self-identify as low-income?"
df["Do you self-identify as low-income? (0-3)"] = df[income_title].apply(lambda x: income_dict[x] if x in income_dict else default_income_score)

school_district_dict = load_values_as_dict("Indicators", school_district_cells)
school_district_title = "What county is your school located in?"
df["School District: DC/PG (0-1)"] = df[school_district_title].apply(lambda x: school_district_dict[x] if x in school_district_dict else default_school_district_score)

ethnicity_race_dict = load_values_as_dict("Indicators", ethnicity_race_cells)
ethnicity_race_title = "Student's Race/Ethnicity"
df["Race/Ethnicity (0-1)"] = df[ethnicity_race_title].apply(lambda x: ethnicity_race_dict[x] if x in ethnicity_race_dict else default_ethnicity_race_score)

gender_dict = load_values_as_dict("Indicators", gender_cells)
gender_title = "Gender Identity (Select all that apply)"

lgbtq_identifying_dict = load_values_as_dict("Indicators", lgbtq_identifying_cells)
lgbtq_title = "I identify as LGBTQIA+"

trans_identifying_dict = load_values_as_dict("Indicators", trans_identifying_cells)
trans_title = "I identify as Transgender, Non-Binary, or Two Spirit"

df["Gender (0-3)"] = df.apply(lambda x: calculate_gender_score(x[gender_title], x[lgbtq_title], x[trans_title]), axis=1)

gpa_dict = load_values_as_dict("Indicators", gpa_cells)
gpa_title = "GPA as of Fall 2022 on a 4.0 scale "

df["GPA (1-4)"] = df[gpa_title].apply(lambda x: gpa_dict[x] if x in gpa_dict else default_gpa_score)

def exp_dict_generator(exp_array):
    d = dict()
    d["No Experience"] = float(exp_array[0])
    d["Somewhat Experienced"] = float(exp_array[1])
    d["Experienced"] = float(exp_array[2])
    d["Very Experienced"] = float(exp_array[3])
    return d

def grades_dict_generator(grades_array):
    d = dict()
    d["A"] = float(grades_array[0])
    d["B"] = float(grades_array[1])
    d["C"] = float(grades_array[2])
    d["D"] = float(grades_array[3])
    return d


c_dict = exp_dict_generator(load_values_as_array(c_exp_cells))
c_title = "Prior Computing Experience. What hands-on experience do you have?  [C# or C++]"
df[c_title] = df[c_title].apply(lambda x: c_dict[x])

java_dict = exp_dict_generator(load_values_as_array(java_exp_cells))
java_title = "Prior Computing Experience. What hands-on experience do you have?  [Java]"
df[java_title] = df[java_title].apply(lambda x: java_dict[x])

js_dict = exp_dict_generator(load_values_as_array(js_exp_cells))
js_title = "Prior Computing Experience. What hands-on experience do you have?  [JavaScript]"
df[js_title] = df[js_title].apply(lambda x: js_dict[x])

py_dict = exp_dict_generator(load_values_as_array(py_exp_cells))
py_title = "Prior Computing Experience. What hands-on experience do you have?  [Python]"
df[py_title] = df[py_title].apply(lambda x: py_dict[x])

html_dict = exp_dict_generator(load_values_as_array(html_exp_cells))
html_title = "Prior Computing Experience. What hands-on experience do you have?  [HTML/CSS]"
df[html_title] = df[html_title].apply(lambda x: html_dict[x])

cyber_dict = exp_dict_generator(load_values_as_array(cyber_exp_cells))
cyber_title = "Prior Computing Experience. What hands-on experience do you have?  [Cybersecurity ]"
df[cyber_title] = df[cyber_title].apply(lambda x: cyber_dict[x])

robot_dict = exp_dict_generator(load_values_as_array(robot_exp_cells))
robot_title = "Prior Computing Experience. What hands-on experience do you have?  [Building Robots/Programming Robots]"
df[robot_title] = df[robot_title].apply(lambda x: robot_dict[x])

gamedev_dict = exp_dict_generator(load_values_as_array(gamedev_exp_cells))
gamedev_title = "Prior Computing Experience. What hands-on experience do you have?  [Game Development]"
df[gamedev_title] = df[gamedev_title].apply(lambda x: gamedev_dict[x])

ai_dict = exp_dict_generator(load_values_as_array(ai_exp_cells))
ai_title = "Prior Computing Experience. What hands-on experience do you have?  [Artificial Intelligence]"
df[ai_title] = df[ai_title].apply(lambda x: ai_dict[x])

hardware_dict = exp_dict_generator(load_values_as_array(hardware_exp_cells))
hardware_title = "Prior Computing Experience. What hands-on experience do you have?  [Hardware/Arduino]"
df[hardware_title] = df[hardware_title].apply(lambda x: hardware_dict[x])

vr_dict = exp_dict_generator(load_values_as_array(vr_exp_cells))
vr_title = "Prior Computing Experience. What hands-on experience do you have?  [Virtual Reality]"
df[vr_title] = df[vr_title].apply(lambda x: vr_dict[x])

scratch_dict = exp_dict_generator(load_values_as_array(scratch_exp_cells))
scratch_title = "Prior Computing Experience. What hands-on experience do you have?  [Block Based/ Drag& Drop Programming (Scratch)]"
df[scratch_title] = df[scratch_title].apply(lambda x: scratch_dict[x])

google_dict = exp_dict_generator(load_values_as_array(google_exp_cells))
google_title = "Prior Computing Experience. What hands-on experience do you have?  [Microsoft Office/Google Apps Suite]"
df[google_title] = df[google_title].apply(lambda x: google_dict[x])

exp_column_names = [c_title, java_title, js_title, py_title, html_title, cyber_title, robot_title, gamedev_title, ai_title, hardware_title, vr_title, scratch_title, google_title]
df["Experience Score (0-100)"] = df[exp_column_names].sum(axis=1)
# df["Experience Score (0-100)"] = df["Experience Score (0-100)"].apply(lambda x: int(str(x)[:-2]) if str(x)[-2:] == ".0" else x)

# science_grade_dict = grades_dict_generator(load_values_as_array(science_grade_cells))
# math_grade_dict = grades_dict_generator(load_values_as_array(math_grade_cells))
# df["Math/Science Grades (1-4)"] = df.apply(lambda x: calculate_grade_score(x["Science Grade"], x["Math Grade"], axis=1))



columns_to_drop = {
    "Prior Computing Experience. What hands-on experience do you have?  [Building Robots/Programming Robots]",
    "Prior Computing Experience. What hands-on experience do you have?  [Block Based/ Drag& Drop Programming (Scratch)]",
    "Prior Computing Experience. What hands-on experience do you have?  [Game Development]",
    "Prior Computing Experience. What hands-on experience do you have?  [HTML/CSS]",
    "Prior Computing Experience. What hands-on experience do you have?  [Microsoft Office/Google Apps Suite]",
    "Prior Computing Experience. What hands-on experience do you have?  [Cybersecurity ]",
    "Prior Computing Experience. What hands-on experience do you have?  [Artificial Intelligence]",
    "Prior Computing Experience. What hands-on experience do you have?  [Virtual Reality]",
    "Prior Computing Experience. What hands-on experience do you have?  [Hardware/Arduino]",
    "Prior Computing Experience. What hands-on experience do you have?  [Java]",
    "Prior Computing Experience. What hands-on experience do you have?  [Python]",
    "Prior Computing Experience. What hands-on experience do you have?  [JavaScript]",
    "Prior Computing Experience. What hands-on experience do you have?  [C# or C++]",
    "GPA as of Fall 2022 on a 4.0 scale ",
    "Gender Identity (Select all that apply)",
    "I identify as LGBTQIA+",
    "I identify as Transgender, Non-Binary, or Two Spirit",
    "Student's Race/Ethnicity",
    "Do you self-identify as low-income?",
    "What is the highest level of education any of your parents has achieved?",
    "What county is your school located in?",
}

df.drop(columns_to_drop, axis="columns", inplace = True)

df["Student ID"] = df.reset_index().index

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

new_column_order = [
    "Student ID",
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
    "Student First Name",
    "Student Last Name",
    "Gender (0-3)",
    "Race/Ethnicity (0-1)",
    "Parent's highest level of education (0-2)",
    "Do you self-identify as low-income? (0-3)",
    "GPA (1-4)",
    "School District: DC/PG (0-1)",
    "Experience Score (0-100)",
    "Current Grade Level",
    "Programs Applied For",
    "CREATE TECH ONLY: How can participating in this camp impact your future education plans?",
    "CYBER DEFENSE ONLY: Please list additional cyber experience (including any training, skills, proficiency, networking, previous cyber camps etc.) ",
    "CYBER DEFENSE ONLY: What are the most important issues that you can address through advanced knowledge of cybersecurity? ",
    "AI4ALL ONLY: How do you think new AI technologies can solve human problems? ",
    "Why you want to attend the program.",
    "What interests you about computing?",
    "Solve a problem using technology",
    "Challenges with technological advances in society",
    "Future career & educational goals.",
    "Relevant middle/high school math classes",
    "Middle/high school technology classes",
    "Highest math class offered at school",
    "Hobbies & Extracurriculars",
    "Prior participation in our programs",
    "Use this space to share any links to your work. (optional)",
    "Use this space to share any additional information with us.",
    "Name of 1st Reference",
    "Name of 2nd Reference"
]

df = df[new_column_order]

print(df)



scopes = ['https://www.googleapis.com/auth/spreadsheets',
          'https://www.googleapis.com/auth/drive']

path = 'i4c-automation-ddbaa5c8fcdd.json'
credentials = Credentials.from_service_account_file(path, scopes=scopes)

gc = gspread.authorize(credentials)

gauth = GoogleAuth()
drive = GoogleDrive(gauth)

# open a google sheet
gs = gc.open_by_key(SPREADSHEET_ID)
# select a work sheet from its name
worksheet1 = gs.worksheet('Output')

worksheet1.clear()
set_with_dataframe(worksheet=worksheet1, dataframe=df, include_index=False,
include_column_header=True, resize=True)
