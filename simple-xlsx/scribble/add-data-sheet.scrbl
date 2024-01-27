#lang scribble/manual

@title{Add Data Sheet}

In the write-xlsx, you can add data sheets.

@section{add-data-sheet}

@codeblock|{
add-data-sheet (-> string? (listof list?) void?)
}|

first argument is sheet name: sheet name is unique, can't have duplicated sheet names.

second argument is a listof list.

data type you can use: string, number, date.

data is a listof list, each list'length can be different.

according to the longest list, function will pad value on the right to keep all the list have the same length.

default pad value is "", you can use #:fill? to specify other values.

ie:
@codeblock|{
(add-data-sheet
  "Sheet1"
  '(
    ("a" "b" "c")
    (1)
    (1.0 2.0)
    ))

 will add the data list below:
 '(
   ("a" "b" "c")
   (1 "" "")
   (1.0 2.0 "")
   )
}|

use #:start_cell? to specify the datalist's start cell, default is "A1".

combine write-xlsx and add-data-sheet, you can generate a xlsx file:

@codeblock|{
(write-xlsx
  basic_file
  (lambda ()
    (add-data-sheet
      "Sheet1"
      '(
        ("month1" "month2" "month3" "month4" "real")
        (201601 100 110 1110 6.9))
       )
     ))
}|
