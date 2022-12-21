// Fill in the spreadsheet ID here
var spreadsheet = SpreadsheetApp.openById("1r9P_cR3RFIuXUMw82wRkWQ_0fvxzSdCy1GSS0PGr8g8");
var data_analysis_sheet = spreadsheet.getSheetByName("Data Analysis"); // Contains the values for each indicator
var source_data_sheet = spreadsheet.getSheetByName("All Data"); // The mass list of all data for all students
var form_response_sheet = spreadsheet.getSheetByName("Form Responses 1"); // Original response data
var upload_sheet = spreadsheet.getSheetByName("Automated Data"); // Where data gets uploaded after processing

var gender_data_set = {};
var ethnicity_race_data_set = {};
var math_science_grades_data_set = {};
var non_math_science_grades_data_set = {};
var school_district_data_set = {};

var default_gender_score = 1;
var default_ethnicity_race_score = 1;
var default_school_district_score = 0;

function response_to_source() {
  var row;

  // For titles
  var current = form_response_sheet.getRange(1, 1, 1, form_response_sheet.getLastColumn()).getValues();
  source_data_sheet.getRange(1, 2, 1, current[0].length).setValues(current);
  var index = source_data_sheet.getLastColumn();
  source_data_sheet.getRange(1, index+1).setValue("Science Grade");
  source_data_sheet.getRange(1, index+2).setValue("Math Grade");
  source_data_sheet.getRange(1, index+3).setValue("Number of A's");
  source_data_sheet.getRange(1, index+4).setValue("Number of B's");
  source_data_sheet.getRange(1, index+5).setValue("Number of C's");
  source_data_sheet.getRange(1, index+5).setValue("Number of D's or Lower");
  

  // For all other rows, we add a student ID
  for (row=2; row <= form_response_sheet.getLastRow(); row++) {
    var current = form_response_sheet.getRange(row, 1, 1, form_response_sheet.getLastColumn()).getValues();
    source_data_sheet.getRange(row, 1, 1, 1).setValue(row); // Student ID num
    source_data_sheet.getRange(row, 2, 1, current[0].length).setValues(current);
  }
}

function source_to_upload_sheet() {
  init_gender_data_set(gender_data_set);
  init_ethnicity_race_data_set(ethnicity_race_data_set);
  init_math_science_grades_data_set(math_science_grades_data_set);
  init_non_math_science_grades_data_set(non_math_science_grades_data_set);
  init_school_district_data_set(school_district_data_set);

  var row;
  for (row=2; row <= source_data_sheet.getLastRow(); row++) {
    process_data(row);
  }


}

function auto_upload_row() {
  init_gender_data_set(gender_data_set);
  init_ethnicity_race_data_set(ethnicity_race_data_set);
  init_math_science_grades_data_set(math_science_grades_data_set);
  init_non_math_science_grades_data_set(non_math_science_grades_data_set);
  init_school_district_data_set(school_district_data_set);

  var last_row = form_response_sheet.getRange(form_response_sheet.getLastRow(), 1, 1, form_response_sheet.getLastColumn()).getValues();

  source_data_sheet.getRange(source_data_sheet.getLastRow()+1, 1).setValue(source_data_sheet.getLastRow()+1);
  source_data_sheet.getRange(source_data_sheet.getLastRow(), 2, 1, last_row[0].length).setValues(last_row);
  process_data(source_data_sheet.getLastRow());


}

function init_gender_data_set(gender_data_set) {
  var gender_data_range = data_analysis_sheet.getRange("A13:B17");
  var r;
  for (r = 1; r < gender_data_range.getNumRows()+1; r++) {
    var key = (gender_data_range.getCell(r, 1).getValue().toString());
    var val = gender_data_range.getCell(r, 2).getValue();
    gender_data_set[key] = val;
  }

}

function gender_calculator(row) {
  var gender_score;
  var gender_column = 22;
  var cell_contents = source_data_sheet.getRange(row, gender_column).getCell(1,1).getValues();

  gender_score = gender_data_set[cell_contents[0]];

  if (gender_score == undefined) {
    gender_score = default_gender_score;
  }

  return gender_score;

}

function test_gender_data_set() {
  console.log(gender_data_set);
  init_gender_data_set(gender_data_set);
  console.log(gender_data_set);
}

function init_ethnicity_race_data_set(ethnicity_race_data_set) {
  var ethnicity_race_range = data_analysis_sheet.getRange("A2:B9");
  var r;
  for (r = 1; r < ethnicity_race_range.getNumRows()+1; r++) {
    var key = (ethnicity_race_range.getCell(r, 1).getValue().toString());
    var val = ethnicity_race_range.getCell(r, 2).getValue();
    ethnicity_race_data_set[key] = val;
  }

}

function ethnicity_race_calculator(row) {
  var ethnicity_race_score;
  var ethnicity_race_column = 23;
  var cell_contents = source_data_sheet.getRange(row, ethnicity_race_column).getCell(1,1).getValues();

  ethnicity_race_score = ethnicity_race_data_set[cell_contents[0]];

  if (ethnicity_race_score==undefined) {
    ethnicity_race_score = default_ethnicity_race_score;
  }

  return ethnicity_race_score;
}

function test_ethnicity_race_data_set() {
  console.log(ethnicity_race_data_set);
  init_ethnicity_race_data_set(ethnicity_race_data_set);
  console.log(ethnicity_race_data_set);
}

function init_math_science_grades_data_set(math_science_grades_data_set) {
  var math_science_grades_range = data_analysis_sheet.getRange("D2:H3");
  var r;
  for (r = 1; r < math_science_grades_range.getNumRows()+1; r++) {
    var key = (math_science_grades_range.getCell(r, 1).getValue().toString());
    var val = {};
      val ["A"] = math_science_grades_range.getCell(r, 2).getValue();
      val ["B"] = math_science_grades_range.getCell(r, 3).getValue();
      val ["C"] = math_science_grades_range.getCell(r, 4).getValue();
      val ["D or Less"] = math_science_grades_range.getCell(r, 5).getValue();
      val [""] = 0;
    math_science_grades_data_set[key] = val;
  }

}

function math_science_grades_calculator(row) {
  var grades_start_col = 35;
  var titles = source_data_sheet.getRange(1, grades_start_col, 1, 2).getValues();
  var student = source_data_sheet.getRange(row, grades_start_col, 1, 2).getValues();

  math_science_sum = 0;
  var index;
  for (index = 0; index < titles[0].length; index++) {
    var category = titles[0][index];
    var math_science_level = student[0][index];
    math_science_sum += math_science_grades_data_set[category][math_science_level];
  }

  if (math_science_sum == 0)
    return "Must manually input report card data";
  else 
    var averaged = math_science_sum/(2);
    averaged = +averaged.toFixed(1); // Rounds to a single decimal point
    return(averaged);

}

function test_math_science_grades_data_set() {
  console.log(math_science_grades_data_set);
  init_math_science_grades_data_set(math_science_grades_data_set);
  console.log(math_science_grades_data_set);
}

function init_non_math_science_grades_data_set(non_math_science_grades_data_set) {
  var non_math_science_grades_range = data_analysis_sheet.getRange("D7:E10");
  var r;
  for (r = 1; r < non_math_science_grades_range.getNumRows()+1; r++) {
    var key = (non_math_science_grades_range.getCell(r, 1).getValue().toString());
    var val = non_math_science_grades_range.getCell(r, 2).getValue();
    non_math_science_grades_data_set[key] = val;
  }

}

function non_math_science_grades_calculator(row) {
  var grades_start_col = 37;

  var student = source_data_sheet.getRange(row, grades_start_col, 1, 4).getValues();

  var grade_sum = 0;
  var num_As = student[0][0];
  var num_Bs = student[0][1];
  var num_Cs = student[0][2];
  var num_Ds_lower = student[0][3];

  grade_sum += (num_As) * (non_math_science_grades_data_set["A"]);
  grade_sum += (num_Bs) * (non_math_science_grades_data_set["B"]);
  grade_sum += (num_Cs) * (non_math_science_grades_data_set["C"]);
  grade_sum += (num_Ds_lower) * (non_math_science_grades_data_set["D or Less"]);


  if (grade_sum == 0)
    return "Must manually input report card data";
  else 
    var averaged = grade_sum/(num_As + num_Bs + num_Cs + num_Ds_lower);
    averaged = +averaged.toFixed(1); // Rounds to a single decimal point
    return(averaged);

}

function test_non_math_science_grades_data_set() {
  console.log(non_math_science_grades_data_set);
  init_non_math_science_grades_data_set(non_math_science_grades_data_set);
  console.log(non_math_science_grades_data_set);
}

function init_school_district_data_set(school_district_data_set) {
  var school_district_data_range = data_analysis_sheet.getRange("G7:H9");
  var r;
  for (r = 1; r < school_district_data_range.getNumRows()+1; r++) {
    var key = (school_district_data_range.getCell(r, 1).getValue().toString());
    var val = school_district_data_range.getCell(r, 2).getValue();
    school_district_data_set[key] = val;
  }

}

function school_district_calculator(row) {
  var school_district_column = 26;
  var school_district_score;
  var cell_contents = source_data_sheet.getRange(row, school_district_column).getCell(1,1).getValues();

  school_district_score = school_district_data_set[cell_contents[0]];

  if (school_district_score == undefined) {
    school_district_score = default_school_district_score;
  }

  return school_district_score;

}

function test_school_district_data_set() {
  console.log(school_district_data_set);
  init_school_district_data_set(school_district_data_set);
  console.log(school_district_data_set);
}

function process_data(row) {
  var first_name_column = 4;
  var last_name_column = 5;
  var grade_level_column = 21;
  var program_column = 13;
  var first_essay_question_column = 14;
  var number_of_essay_questions = 6;
  // var first_reference_name_column = 35;
  // var second_reference_name_column = 36;

  // The following are the column numbers into which the data will be uploaded
  var student_id_col = 1;
  var first_name_col = 12;
  var last_name_col = 13;
  var gender_score_col = 14;
  var ethnicity_race_score_col = 15;
  var math_science_grades_col = 16;
  var non_math_science_grades_col = 17;
  var school_district_col = 18;
  var grade_level_col = 21;
  var program_name_col = 22;
  var auto_fill_data_start_col = 23;

  var data_row = source_data_sheet.getRange(row, 1, 1, source_data_sheet.getLastColumn());
  var upload_location = upload_sheet;

  upload_location.getRange(row, student_id_col).setValue(row); // Student Id (same as row number)
  upload_location.getRange(row, first_name_col).setValue(data_row.getCell(1, first_name_column).getValue()); // First Name
  upload_location.getRange(row, last_name_col).setValue(data_row.getCell(1, last_name_column).getValue()); // Last Name
  upload_location.getRange(row, gender_score_col).setValue(gender_calculator(row)); // Gender
  upload_location.getRange(row, ethnicity_race_score_col).setValue(ethnicity_race_calculator(row)); // Race/Ethnicity
  upload_location.getRange(row, math_science_grades_col).setValue(math_science_grades_calculator(row)); // Math and Science Grades
  upload_location.getRange(row, non_math_science_grades_col).setValue(non_math_science_grades_calculator(row)); // All Other Grades
  upload_location.getRange(row, school_district_col).setValue(school_district_calculator(row)); // School District
  upload_location.getRange(row, grade_level_col).setValue(data_row.getCell(1, grade_level_column).getValue()); // What grade they're currently in
  upload_location.getRange(row, program_name_col).setValue(data_row.getCell(1, program_column).getValue()); // Which Program they're applying for
  var col;
  for (col = auto_fill_data_start_col; col < number_of_essay_questions+auto_fill_data_start_col; col++) {
    upload_location.getRange(row,col).setValue(data_row.getCell(1, first_essay_question_column).getValue());
    first_essay_question_column += 1;
  }

  // upload_location.getRange(row,col).setValue(data_row.getCell(1, first_reference_name_column).getValue()); // Name of First Reference
  // upload_location.getRange(row,col+1).setValue(data_row.getCell(1, second_reference_name_column).getValue()); // Name of Second Reference
    
}

function test_process_data() {
  init_gender_data_set(gender_data_set);
  init_ethnicity_race_data_set(ethnicity_race_data_set);
  init_math_science_grades_data_set(math_science_grades_data_set);
  init_non_math_science_grades_data_set(non_math_science_grades_data_set);
  init_school_district_data_set(school_district_data_set);

  process_data(2);
}





