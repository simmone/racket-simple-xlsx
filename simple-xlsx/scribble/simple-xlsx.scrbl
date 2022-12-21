#lang scribble/manual

@title{Simple-Xlsx: Handle .xlsx in Racket}

@author+email["Chen Xiao" "chenxiao770117@gmail.com"]

The @tt{simple-xlsx} read and write @tt{.xlsx} file.

1. Support string, number, date as data type.

2. Multiple style supported: color, border, font, etc.

3. Link data sheet to generate chart sheet.

4. Read and write combine one process: read, modified something, write back, all in a lambda scope.

@table-of-contents[]

@section[#:tag "install"]{Install}

raco pkg install simple-xlsx

Caution: simple-xlsx depends on package: simple-xml, if you have installed simple-xml, should update the package to newest version.

@section{Basic Example}

@subsection{Generate a xlsx file}

@verbatim|{
(write-xlsx
  "basic.xlsx"
  (lambda ()
    (add-data-sheet "Sheet1" '(("month1" "month2" "month3" "month4" "real")))

    (add-data-sheet "Sheet2" '((201601 100 110 1110 6.9)))))
}|

1. All operations in @tt{write-xlsx}'s lambda scope.

2. Specify file name, sheet name, have same count's list data, done.

@subsection{Read a xlsx file}

@verbatim|{
(read-xlsx
  "basic_write.xlsx"
  (lambda ()
    (check-equal? (get-sheet-name-list) '("Sheet1" "Sheet2"))

    (with-sheet-ref
    0
    (lambda ()
      (check-equal? (get-row 1) '("month1" "month2" "month3" "month4" "real"))))

    (with-sheet-ref
    1
    (lambda ()
      (check-equal? (get-row 1) '(201601 100 110 1110 6.9))))))
}|

@subsection{Read a xlsx file, modify something and Write back}

@verbatim|{
(read-and-write-xlsx
  basic_write_file
  basic_read_and_write_file
  (lambda ()
    (check-equal? (get-sheet-name-list) '("Sheet1" "Sheet2"))

    (with-sheet-ref
    0
    (lambda ()
      (set-cell-value! "B1" "John")
      (check-equal? (get-row 1) '("month1" "John" "month3" "month4" "real"))))

    (with-sheet-ref
    1
    (lambda ()
      (check-equal? (get-row 1) '(201601 100 110 1110 6.9))))
      ))
}|

The first arg is read file, second is the write back file, these two can be a same file, if you want to replace the oringinal.

@include-section["xlsx.scrbl"]

@include-section["sheet.scrbl"]

@include-section["add-data-sheet.scrbl"]

@include-section["access-data.scrbl"]

@include-section["style/style.scrbl"]

@include-section["chartsheet/chartsheet.scrbl"]

