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

caution: each list of data's count must be same.

(add-data-sheet "Sheet1" '((1 2 3) ("1" "a" 5))) is ok.

(add-data-sheet "Sheet1" '((1 2 3 4) ("1" "a" 7 8)) is error.

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
