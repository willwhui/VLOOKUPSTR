# VLOOKUPSTR
Note:
It's a google spreadsheet add-on:
https://chrome.google.com/webstore/detail/vlookupstr/igkbgmafodaojojegoijmffeogjfnpda

Source code:
vlookupstr.js

Discription on chrome store:

This function performs a none-full-match string searching, which can not be done by VLOOKUP.
VLOOKUPSTR(search_key, range, index, case_insensitive)

This function acts like the build-in function VLOOKUP. But it performs a non-full-match string searching, which can not be done by VLOOKUP.

Example:
=VLOOKUPSTR(A1, A2:B7, 2, FALSE)
To Check for the string of the  cell A1 in the column A from A2:B7, if any cell contains ( not equal to) the string, for example, A1= "File" and A6="File Menu", then the function returns the cell value of B6 in this example.

And the differences are: 
1. as above, if the search_key is the substring of the cell in the first column, then the function will return as found. 
2. can perform a none case insensitive search 
3. always assume the column to be searched is not sorted

search_key
The value to search for.

range
The range to consider for the search.

index
The column index of the value to be returned.

case_insensitive
Indicates whether the searching is case insensitive.
