#lang scribble/manual

@title{Second Level: with-sheet*}

In the with-sheet-*'s scope, you can get and set sheet data, set sheet styles, etc.

@section{with-sheet-*}

@codeblock|{
with-sheet-ref (-> natural? procedure? any)
with-sheet (-> procedure? any)
with-sheet-name (-> string? procedure? any)
}|

with-sheet-ref's first argument is the sheet index, start from 0.
with-sheet means with-sheet-ref 0.
with-sheet-name use sheet name to specify.

All the sheet data's operations: get or set data, set data's styles should be placed in the with-sheet-*.

Because most methods is effect in the sheet scope, so normally, the code style is like below:

@codeblock|{
(write-xlsx
  "out.xlsx"
  (lambda ()
    (add-data-sheet "Sheet1" '(("month1" "month2" "month3" "month4" "real")))

    (with-sheet-ref
    0
    (lambda ()
      (set-cell! "B1" "John")))))
}|
