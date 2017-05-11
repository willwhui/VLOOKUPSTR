/**
Open source license: LGPL
Author: willwhui@gmail.com 
*/

var ERR_MSG_NA = "#VALUE_NOT_FOUND"
var ERR_MSG_OUT_RANGE = "#INDEX_OUT_OF_RANGE"
var ERR_MSG_VALUE_ARRAY = "#VALUE_ARRAY_NOT_ALLOWED"
var range_;
var index_;
var case_;

/**
 * Acts mostly like "vlookup", the diffrences are:
 * 1. if the search_key is the sub-string of the cell in the first column, then the function will return as found.
 * 2. can perform a none case insensitive search
 * 3. always assume the column to be searched is not sorted
 * @param {A1} search_key The value to search for.
 * @param {A2:B26} range The range to consider for the search.
 * @param {2} index The column index of the value to be returned.
 * @param {FALSE} case_insensitive Indicates whether the searching is case insensitive.
 * @return The value of the specified cell in the row found.
 * @customfunction
 */
function VLOOKUPSTR(search_key, range, index, case_insensitive) {
  range_ = range;
  index_ = index;
  case_ = Number(case_insensitive);

  // index checking
  if (index <= 0) {
    return ERR_MSG_OUT_RANGE;
  }
  if (range instanceof Array) { // array
    subRange = range[0];
    if (subRange instanceof Array) { // 2D array
      if (index > subRange.length) {
        return ERR_MSG_OUT_RANGE;
      }
    }
    else { // 1D array
      if (index != 1) {
        return ERR_MSG_OUT_RANGE;
      }
    }
  }
  else { // single value
    if (index != 1) {
      return ERR_MSG_OUT_RANGE;
    }
  }
  // convert key to string once for the whole process
  key = array_to_string(search_key);

  // convert range first column to string for the whole process
  range_first_col_to_string();
  
  return vlookupstr_(key);
}

function vlookupstr_(key) {
  if (key.map) {
    return key.map(vlookupstr_); // extrac key to single value
  }
  else {
    return vlookupstr__(key, range_, index_, case_);
  }
}

function vlookupstr__(key, range, index, case_insensitive) {  
  // searching
  if (!(range instanceof Array)) {
    return vls_in_value_(key, range, index, case_insensitive);
  }
  else {
    return vls_in_cols_(key, range, index, case_insensitive);
  }
}

function vls_in_value_(key, value, index, case_insensitive) {
  if (index > 1) {
    return ERR_MSG_OUT_RANGE;
  }
  if (-1 != value.indexOf(key)) {
    return value;
  }
  else {
    return ERR_MSG_NA;
  }
}

function vls_in_cols_(key, range, index, case_insensitive) {
  arrayLength = range.length;  
  for (var i = 0; i < arrayLength; i++) {
    if (-1 != range[i][0].indexOf(key)) {
      return range[i][index - 1];
    }
  }
  return ERR_MSG_NA;
}

function array_to_string(input) {
  if (input.map) {            // Test whether input is an array.
    return input.map(array_to_string); // Recurse over array if so.
  }
  else {
    return trans_to_string_(input, case_);
  }
}

function range_first_col_to_string() {
  if (range_ instanceof Array) {
    len = range_.length;
    for (i = 0; i < len; i++) {
      range_[i][0] = trans_to_string_(range_[i][0], case_);
    }
  }
  else {
    range_[0] = trans_to_string_(range_[0], case_);
  }
}

function trans_to_string_(p, case_insensitive) {
  if (!(p instanceof String)) {
    p = p.toString();
  }
  if (!case_insensitive) {
    p = p.toLowerCase();
  }
  return p;
}

function onInstall(e) {
  onOpen(e);
}

function onOpen(e) {
  var menu = SpreadsheetApp.getUi().createAddonMenu(); // Or DocumentApp or FormApp.
  if (e && e.authMode == ScriptApp.AuthMode.NONE) {
    // Add a normal menu item (works in all authorization modes).
    menu.addItem('How to use', 'menuHowToUseCallback');
  } else {
      menu.addItem('How to use', 'menuHowToUseCallback');
  }
  menu.addToUi();
}

var SCRIPTION = 'VLOOKUPSTR(search_key, range, index, case_insensitive)\
\n\
\n\
This function acts like the build-in function VLOOKUP. But it performs a non-full-match string searching, which can not be done by VLOOKUP.\
\n\
\n\
Example:\
\n\
=VLOOKUPSTR(A1, A2:B26, 2, FALSE)\
\n\
To Check for the string of the  cell A1 in the column A from A2:B26, if any cell contains ( not equal to) the string, for example, A1= "File" and A6="File Menu", then the function returns the cell value of B6 in this example.\
\n\
\n\
And the differences are: \
\n\
1. as above, if the search_key is the substring of the cell in the first column, then the function will return as found. \
\n\
2. can perform a none case insensitive search \
\n\
3. always assume the column to be searched is not sorted\
\n\
\n\
search_key\
\n\
The value to search for.\
\n\
\n\
range\
\n\
The range to consider for the search.\
\n\
\n\
index\
\n\
The column index of the value to be returned.\
\n\
\n\
case_insensitive\
\n\
Indicates whether the searching is case insensitive.\
'

function menuHowToUseCallback() {
  var ui = SpreadsheetApp.getUi();
ui.alert('How to use', SCRIPTION, ui.ButtonSet.OK);
}
