// Fill in the spreadsheet ID here
var spreadsheet = SpreadsheetApp.openById("1TnwuYCViBkKxFFkockuDx7Dp9fe-GZp0cfvEQs3mFD4");
var data_analysis_sheet = spreadsheet.getSheetByName("Data Analysis"); // Contains the values for each indicator
var source_data_sheet = spreadsheet.getSheetByName("All Data"); // The mass list of all data for all students
var form_response_sheet = spreadsheet.getSheetByName("Form Responses 1"); // Original response data
var upload_sheet = spreadsheet.getSheetByName("Automated Data"); // Where data gets uploaded after processing

var gender_data_set = {};
var ethnicity_race_data_set = {};
var parent_education_data_set = {};
var gpa_data_set = {};
var math_science_grades_data_set = {};
var school_district_data_set = {};
var exp_data_set = {};
var orientation_data_set = {};
var income_data_set = {};
var reviewer_data_set = [];


// Update the cell numbers here from the Data Analysis Tab
var gender_cells = "A19:B24";
var ethnicity_race_cells = "G2:H12";
var parent_education_cells = "G16:H27";
var gpa_cells = "D27:E33";
var math_science_grades_cells = "J2:N3";
var school_district_cells = "D19:E21";
var exp_cells = "A2:E14";
var orientation_cells = "J8:N9";
var income_cells = "G31:H34";
var reviewer_cells = "J15:J23";

// Update the column numbers here from the All Data tab
var first_name_location = 3;
var last_name_location = 4;
var gender_score_location = 48;
var ethnicity_race_score_location = 52;
var parent_education_location = 54;
var gpa_location = 45;
var math_science_grades_location = 72;
var school_district_location = 56;
var experience_score_location = 26;
var grade_level_location = 47;
var program_name_location = 15;
var program_specific_q1_location = 39;
var program_specific_q2_location = 40;
var program_specific_q3_location = 41;
var program_specific_q4_location = 42;
var auto_fill_data_start_location = 27;
var work_links_location = 43;
var additional_info_location = 44;
var reference1_location = 63;
var reference2_location = 66;
var school_district_location = 56;
var parent_education_location = 54;
var gender_location = 48;
var ethnicity_race_location = 52;
var gpa_location = 45;
var orientation_location = 49;
var income_location = 53;

var default_gender_score = 1;
var default_ethnicity_race_score = 1;
var default_parent_education_score = 0;
var default_gpa_score = 0;
var default_school_district_score = 0;
var default_income_score = 0;

function test_all_data_sets() {
  test_data_set(school_district_data_set);
  test_data_set(math_science_grades_data_set);
  test_data_set(ethnicity_race_data_set);
  test_data_set(gender_data_set);
  test_data_set(exp_data_set);
  test_data_set(gpa_data_set);
  test_data_set(parent_education_data_set);
  test_data_set(orientation_data_set);
  test_data_set(income_data_set);
  test_data_set(reviewer_data_set);
}


function response_to_source() {
  var row;
  var header = form_response_sheet.getRange(1, 1, 1, form_response_sheet.getLastColumn()).getValues();
  source_data_sheet.getRange(1, 1, 1, 1).setValue("Student ID");
  source_data_sheet.getRange(1, 2, 1, header[0].length).setValues(header);
  source_data_sheet.getRange(1, form_response_sheet.getLastColumn()+2, 1, 1).setValue("Science Grade");
  source_data_sheet.getRange(1, form_response_sheet.getLastColumn()+3, 1, 1).setValue("Math Grade");
  source_data_sheet.getRange(1, form_response_sheet.getLastColumn()+4, 1, 1).setValue("Notes");
  for (row=2; row <= form_response_sheet.getLastRow(); row++) {
    var current = form_response_sheet.getRange(row, 1, 1, form_response_sheet.getLastColumn()).getValues();
    source_data_sheet.getRange(row, 1, 1, 1).setValue(row); // Student ID num
    source_data_sheet.getRange(row, 2, 1, current[0].length).setValues(current);
  }
}

function source_to_upload_sheet_20() {
  init_all_sets();

  var row;
  for (row=2; row <= 20; row++) {
    process_data(row);
  }

}

function source_to_upload_sheet_40() {
  init_all_sets();

  var row;
  for (row=21; row <= 40; row++) {
    process_data(row);
  }

}

function source_to_upload_sheet_60() {
  init_all_sets();

  var row;
  for (row=41; row <= 60; row++) {
    process_data(row);
  }

}

function source_to_upload_sheet_80() {
  init_all_sets();

  var row;
  for (row=61; row <= 80; row++) {
    process_data(row);
  }

}

function source_to_upload_sheet_100() {
  init_all_sets();

  var row;
  for (row=81; row <= 100; row++) {
    process_data(row);
  }

}

function source_to_upload_sheet_120() {
  init_all_sets();

  var row;
  for (row=101; row <= 120; row++) {
    process_data(row);
  }

}

function source_to_upload_sheet_140() {
  init_all_sets();

  var row;
  for (row=121; row <= 140; row++) {
    process_data(row);
  }

}

function source_to_upload_sheet_160() {
  init_all_sets();

  var row;
  for (row=141; row <= 160; row++) {
    process_data(row);
  }

}

function source_to_upload_sheet_last() {
  init_all_sets();

  var row;
  for (row=161; row <= source_data_sheet.getLastRow(); row++) {
    process_data(row);
  }

}

function auto_upload_row() {
  init_all_sets();

  var last_row = form_response_sheet.getRange(form_response_sheet.getLastRow(), 1, 1, form_response_sheet.getLastColumn()).getValues();

  source_data_sheet.getRange(source_data_sheet.getLastRow()+1, 1).setValue(source_data_sheet.getLastRow()+1);
  source_data_sheet.getRange(source_data_sheet.getLastRow(), 2, 1, last_row[0].length).setValues(last_row);
  process_data(source_data_sheet.getLastRow());


}

// Accounts for both gender and orientation
function gender_calculator(row) {
  var gender_score;
  var cell_contents = source_data_sheet.getRange(row, gender_location).getCell(1,1).getValues();

  gender_score = gender_data_set[cell_contents[0]];

  if (gender_score == undefined) {
    gender_score = default_gender_score;
  }

  var orientation_score = 0;
  var titles = source_data_sheet.getRange(1, orientation_location, 1, 2).getValues();
  var student = source_data_sheet.getRange(row, orientation_location, 1, 2).getValues();

  var index;
  for (index = 0; index < titles[0].length; index++) {
    var category = titles[0][index];
    var answer = student[0][index];
    orientation_score += orientation_data_set[category][answer];
  }

  return(gender_score + orientation_score);

}

function ethnicity_race_calculator(row) {
  var ethnicity_race_score;
  var cell_contents = source_data_sheet.getRange(row, ethnicity_race_location).getCell(1,1).getValues();

  ethnicity_race_score = ethnicity_race_data_set[cell_contents[0]];

  if (ethnicity_race_score==undefined) {
    ethnicity_race_score = default_ethnicity_race_score;
  }

  return ethnicity_race_score;
}

function parent_education_calculator(row) {
  var parent_education_score;
  var cell_contents = source_data_sheet.getRange(row, parent_education_location).getCell(1,1).getValues();

  parent_education_score = parent_education_data_set[cell_contents[0]];

  if (parent_education_score == undefined) {
    parent_education_score = default_parent_education_score;
  }

  return parent_education_score;

}

function income_calculator(row) {
  var income_score;
  var cell_contents = source_data_sheet.getRange(row, income_location).getCell(1,1).getValues();

  income_score = income_data_set[cell_contents[0]];

  if (income_score==undefined) {
    income_score = default_income_score;
  }

  return income_score;
}

function gpa_calculator(row) {
  var gpa_score;
  var cell_contents = source_data_sheet.getRange(row, gpa_location).getCell(1,1).getValues();

  gpa_score = gpa_data_set[cell_contents[0]];

  if (gpa_score == undefined) {
    gpa_score = default_gpa_score;
  }

  return gpa_score;

}

function math_science_grades_calculator(row) {
  var titles = source_data_sheet.getRange(1, math_science_grades_location, 1, 2).getValues();
  var student = source_data_sheet.getRange(row, math_science_grades_location, 1, 2).getValues();

  math_science_sum = 0;
  var index;
  for (index = 0; index < titles[0].length; index++) {
    var category = titles[0][index];
    var math_science_level = student[0][index];
    math_science_sum += math_science_grades_data_set[category][math_science_level];
  }

  if (math_science_sum == 0)
    return "Grades Missing for Student in All Data Sheet";
  else 
    var averaged = math_science_sum/(2);
    averaged = +averaged.toFixed(1); // Rounds to a single decimal point
    return(averaged);

}

function school_district_calculator(row) {
  var school_district_score;
  var cell_contents = source_data_sheet.getRange(row, school_district_location).getCell(1,1).getValues();

  school_district_score = school_district_data_set[cell_contents[0]];

  if (school_district_score == undefined) {
    school_district_score = default_school_district_score;
  }

  return school_district_score;

}

function experience_calculator(row) {
  var titles = source_data_sheet.getRange(1, experience_score_location, 1, 13).getValues();
  var student = source_data_sheet.getRange(row, experience_score_location, 1, 13).getValues();

  exp_sum = 0;
  var index;
  for (index = 0; index < titles[0].length; index++) {
    var category = titles[0][index];
    var exp_level = student[0][index];
    exp_sum += exp_data_set[category][exp_level];
  }

  return exp_sum;

}

function reviewer_calculator(row) {
  init_all_sets();
  index = row % reviewer_data_set.length;
  return reviewer_data_set[index];

}

function process_data(row) {

  // What column the value gets put into
  var student_id_col = 1;
  var reviewer_col = 2;
  var first_name_col = 12;
  var last_name_col = 13;
  var gender_score_col = 14;
  var ethnicity_race_score_col = 15;
  var parent_education_col = 16;
  var income_col = 17;
  var gpa_col = 18;
  var math_science_grades_col = 19;
  var school_district_col = 20;
  var experience_score_col = 21;
  var grade_level_col = 22;
  var program_name_col = 23;
  var program_specific_q1_col = 24;
  var program_specific_q2_col = 25;
  var program_specific_q3_col = 26;
  var program_specific_q4_col = 27;
  var auto_fill_data_start_col = 28;
  var work_links_col = 38;
  var additional_info_col = 39;
  var reference1_col = 40;
  var reference2_col = 41;

  var data_row = source_data_sheet.getRange(row, 1, 1, 68);
  var upload_location = upload_sheet;

  upload_location.getRange(row, student_id_col).setValue(row); // Student Id (same as row number)
  upload_location.getRange(row, reviewer_col).setValue(reviewer_calculator(row)); // Reviewer
  upload_location.getRange(row, first_name_col).setValue(data_row.getCell(1, first_name_location).getValue()); // First Name
  upload_location.getRange(row, last_name_col).setValue(data_row.getCell(1, last_name_location).getValue()); // Last Name
  upload_location.getRange(row, gender_score_col).setValue(gender_calculator(row)); // Gender
  upload_location.getRange(row, ethnicity_race_score_col).setValue(ethnicity_race_calculator(row)); // Race/Ethnicity
  upload_location.getRange(row, parent_education_col).setValue(parent_education_calculator(row)); // Parent Education
  upload_location.getRange(row, income_col).setValue(income_calculator(row)); // Low-Income
  upload_location.getRange(row, gpa_col).setValue(gpa_calculator(row)); // GPA
  upload_location.getRange(row, math_science_grades_col).setValue(math_science_grades_calculator(row)); // Math and Science Grades
  upload_location.getRange(row, school_district_col).setValue(school_district_calculator(row)); // School District
  upload_location.getRange(row, experience_score_col).setValue(experience_calculator(row)); // Experience Score
  upload_location.getRange(row, grade_level_col).setValue(data_row.getCell(1, grade_level_location).getValue()); // What grade they're currently in
  upload_location.getRange(row, program_name_col).setValue(data_row.getCell(1, program_name_location).getValue()); // Which Program they're applying for
  
  upload_location.getRange(row, program_specific_q1_col).setValue(data_row.getCell(1, program_specific_q1_location).getValue());
  upload_location.getRange(row, program_specific_q2_col).setValue(data_row.getCell(1, program_specific_q2_location).getValue());
  upload_location.getRange(row, program_specific_q3_col).setValue(data_row.getCell(1, program_specific_q3_location).getValue());
  upload_location.getRange(row, program_specific_q4_col).setValue(data_row.getCell(1, program_specific_q4_location).getValue());

  var col;
  for (col = auto_fill_data_start_col; col <= 9+auto_fill_data_start_col; col++) {
    upload_location.getRange(row,col).setValue(data_row.getCell(1, col-12).getValue());
  }

  upload_location.getRange(row,work_links_col).setValue(data_row.getCell(1, work_links_location).getValue()); // Links to Work
  upload_location.getRange(row,additional_info_col).setValue(data_row.getCell(1, additional_info_location).getValue()); // Additional Info
  upload_location.getRange(row,reference1_col).setValue(data_row.getCell(1, reference1_location).getValue()); // Name of First Reference
  upload_location.getRange(row,reference2_col).setValue(data_row.getCell(1, reference2_location).getValue()); // Name of Second Reference

}

function test_data_set(set) {
  console.log(set);
  if (set == gender_data_set) {
    init_data_set(set, gender_cells);
  } else if (set == ethnicity_race_data_set) {
    init_data_set(set, ethnicity_race_cells);
  } else if (set == parent_education_data_set) {
    init_data_set(set, parent_education_cells);
  } else if (set == gpa_data_set) {
    init_data_set(set, gpa_cells);
  } else if (set == math_science_grades_data_set) {
    init_data_set(set, math_science_grades_cells);
  } else if (set == school_district_data_set) {
    init_data_set(set, school_district_cells);    
  } else if (set == exp_data_set) {
    init_data_set(set, exp_cells);      
  } else if (set == orientation_data_set) {
    init_data_set(set, orientation_cells);
  } else if (set == income_data_set) {
    init_data_set(set, income_cells);
  } else if (set == reviewer_data_set) {
    init_data_set(set, reviewer_cells);
  }
  console.log(set);
}

function init_data_set(set, cell_range) {
  var data_range = data_analysis_sheet.getRange(cell_range);
  var r;
  if (set == exp_data_set) {
    for (r = 1; r < data_range.getNumRows()+1; r++) {
      var key = (data_range.getCell(r, 1).getValue().toString());
      var val = {};
        val ["No Experience"] = data_range.getCell(r, 2).getValue();
        val ["Somewhat Experienced"] = data_range.getCell(r, 3).getValue();
        val ["Experienced"] = data_range.getCell(r, 4).getValue();
        val ["Very Experienced"] = data_range.getCell(r, 5).getValue();
      set[key] = val;
    }

  } else if (set == math_science_grades_data_set) {
    for (r = 1; r < data_range.getNumRows()+1; r++) {
      var key = (data_range.getCell(r, 1).getValue().toString());
      var val = {};
        val ["A"] = data_range.getCell(r, 2).getValue();
        val ["B"] = data_range.getCell(r, 3).getValue();
        val ["C"] = data_range.getCell(r, 4).getValue();
        val ["D or Less"] = data_range.getCell(r, 5).getValue();
        val [""] = 0;
      set[key] = val;
    }
  } else if (set == orientation_data_set) {
    for (r = 1; r < data_range.getNumRows()+1; r++) {
      var key = (data_range.getCell(r, 1).getValue().toString());
      var val = {};
        val ["Yes"] = data_range.getCell(r, 2).getValue();
        val ["No"] = data_range.getCell(r, 3).getValue();
        val ["Not Sure"] = data_range.getCell(r, 4).getValue();
        val ["Prefer not to disclose"] = data_range.getCell(r, 5).getValue();
      set[key] = val;
    }
  } else if (set == reviewer_data_set) {
    for (r = 1; r < data_range.getNumRows()+1; r++) {
      set.push(data_range.getCell(r, 1).getValue().toString())
    }
  } else {
    for (r = 1; r < data_range.getNumRows()+1; r++) {
      var key = (data_range.getCell(r, 1).getValue().toString());
      var val = data_range.getCell(r, 2).getValue();
      set[key] = val;
    }
  }

}

function init_all_sets() {
  init_data_set(gender_data_set, gender_cells);
  init_data_set(ethnicity_race_data_set, ethnicity_race_cells);
  init_data_set(math_science_grades_data_set, math_science_grades_cells);
  init_data_set(school_district_data_set, school_district_cells);
  init_data_set(parent_education_data_set, parent_education_cells);
  init_data_set(gpa_data_set, gpa_cells);
  init_data_set(exp_data_set, exp_cells);
  init_data_set(orientation_data_set, orientation_cells);
  init_data_set(income_data_set, income_cells);
  init_data_set(reviewer_data_set, reviewer_cells);
}

function test_process_data() {
  init_all_sets();
  process_data(2);
}

function indicator() {
  init_all_sets();
  var row;
  var col = 19;
  for (row=2; row <= upload_sheet.getLastRow(); row++) {
    upload_sheet.getRange(row, col).setValue(math_science_grades_calculator(row));
    console.log(row);
  }
}

function process_all() {
  init_all_sets();

  var row;
  for (row=2; row <= source_data_sheet.getLastRow(); row++) {
    process_data(row);
    console.log(row);
  }

}

function manual() {
  init_all_sets();

  var last_row = form_response_sheet.getRange(form_response_sheet.getLastRow(), 1, 1, form_response_sheet.getLastColumn()).getValues();

  source_data_sheet.getRange(source_data_sheet.getLastRow()+1, 1).setValue(source_data_sheet.getLastRow()+1);
  source_data_sheet.getRange(source_data_sheet.getLastRow(), 2, 1, last_row[0].length).setValues(last_row);
  process_data(source_data_sheet.getLastRow());
}


