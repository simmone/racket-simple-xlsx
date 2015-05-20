#lang scribble/manual

@(require (for-label racket))

@title{simple-xlsx package usage}

@author+email["Chen Xiao" "chenxiao770117@gmail.com"]

simple-xlsx package is a package to read and write .xlsx format file.@linebreak{}
.xlsx file is a open xml format file.

@table-of-contents[]

@section[#:tag "install"]{Install}

raco pkg install simple-xlsx

@section{Read}

read from a .xlsx file.@linebreak{}
you can get a specific cell's value or loop for the whole sheet's rows.

@defmodule[simple-xlsx]
@(require (for-label simple-xlsx))

there is also a complete read and write example on github:@link["https://github.com/simmone/racket-simple-xlsx/blob/master/example/example.rkt"]{includedin the source}.

@defproc[(with-input-from-xlsx-file
              [xlsx_file_path (path-string?)]
              [user-proc (-> xlsx_handler void?)])
            void?]{
  read xlsx's main outer func, all read assosiated action is include in user-proc.
}

@defproc[(load-sheet
           [sheet_name (string?)]
           [xlsx_handler (xlsx_handler)])
           void?]{
  load specified sheet by sheet name.@linebreak{}
  must first called before other func, because any other func is based on specified sheet.
}

@defproc[(get-sheet-names
            [xlsx_handler (xlsx_handler)])
            list?]{
  get sheet names.@linebreak{}
}

@defproc[(get-cell-value
            [cell_axis (string?)]
            [xlsx_handler (xlsx_handler)])
            any]{
  get cell value through cell's axis.@linebreak{}
  cell axis: A1 B2 C3...
}

@defproc[(get-sheet-dimension
            [xlsx_handler (xlsx_handler)])
            string?]{
  get current sheet's dimension.@linebreak{}
  like A1:C5
}

@defproc[(with-row
            [xlsx_handler (xlsx_handler)]
            [user_proc (-> list? any)])
            any]{
  with-row is like for-each, deal a line one time.
}

@section{Write}

write a xlsx file means instance a xlsx-data% class.@linebreak{}
use add-sheet to fill a xlsx-data% instance and write it to file.

@subsection{xlsx-data%}

xlsx-data% class represent a whole xlsx file's data.@linebreak{}
it contains sheet data and col attibutes.@linebreak{}
xlsx-data% have only one method: add-sheet.@linebreak{}
sheet data just a list: (list (list cell ...) (list cell ...)...).@linebreak{}

this is a simple xlsx-data% example without col attributes:@linebreak{}

@verbatim{
  (let ([xlsx (new xlsx-data%)])
    (send xlsx add-sheet '(("chenxiao" "cx") (1 2 34 100 456.34)) "Sheet1")
}

@subsection{col-attribute}

col-attribute used to specify column's attribute.@linebreak{}
you can set column's width, color by now.@linebreak{}

col-attribute is a hash.@linebreak{}
hash's key is column's dimension, value is a struct named:col-attr.@linebreak{}

for example:@linebreak{}
@verbatim{
  ;; set column A width: 100, color: FF0000
  (hash-set! col_attr_hash "A" (col-attr 100 "FF0000"))
  (hash-set! col_attr_hash "B" (col-attr 200 "00FF00"))
  (hash-set! col_attr_hash "C" (col-attr 200 "EF9595"))
  (hash-set! col_attr_hash "D-F" (col-attr 100 "0000FF"))
  (hash-set! col_attr_hash "7-10" (col-attr 100 "EE89CD"))
}

@subsection{func}

@defproc[(write-xlsx-file
            [xlsx-data (xlsx-data%)]
            [path (path-string?)])
            void?]{
  write xlsx-data% to xlsx file.
}

@section{Example}

@verbatim{
  #lang racket

  (require simple-xlsx)

  ;; write a xlsx file, with multiple sheets, set the column attributes(width, color)
  (let ([xlsx (new xlsx-data%)]
        [col_attr_hash (make-hash)])
      (hash-set! col_attr_hash "A" (col-attr 100 "FF0000"))
      (hash-set! col_attr_hash "B" (col-attr 200 "00FF00"))
      (hash-set! col_attr_hash "C" (col-attr 200 "EF9595"))
      (hash-set! col_attr_hash "D-F" (col-attr 100 "0000FF"))
      (hash-set! col_attr_hash "7-10" (col-attr 100 "EE89CD"))
  
      (send xlsx add-sheet '(("Jane Birkin" "Leonard Cohen" "Matthew McConaughey") (1 2 34 100 456.34)) "Sheet1" #:col_attr_hash col_attr_hash)
      (send xlsx add-sheet '((1 2 3 4)) "Sheet2")
      (send xlsx add-sheet '(("a" "b")) "Sheet3")
      (send xlsx add-sheet '(("a" "b")) "Sheet4" #:col_attr_hash col_attr_hash)
      (send xlsx add-sheet '(("")) "Sheet5" #:col_attr_hash col_attr_hash)
      (write-xlsx-file xlsx "test1.xlsx"))
  
  ;; write a xlsx file and read it back
  (let ([xlsx (new xlsx-data%)])
    (send xlsx add-sheet '(("chenxiao" "cx") (1 2 34 100 456.34)) "Sheet1")
    (send xlsx add-sheet '((1 2 3 4)) "Sheet2")
    (send xlsx add-sheet '(("")) "Sheet3")
    (write-xlsx-file xlsx "test2.xlsx")
  
    ;; read specific cell
    (with-input-from-xlsx-file
     "test2.xlsx"
     (lambda (xlsx)
       (printf "~a\n" (get-sheet-names xlsx)) ;(Sheet1 Sheet2 Sheet3)
       
       (load-sheet "Sheet1" xlsx)
  
       (printf "~a\n" (get-sheet-dimension xlsx)) ; (2 . 5)
       (printf "~a\n" (get-cell-value "A1" xlsx)) ; "chenxiao"
       (printf "~a\n" (get-cell-value "B1" xlsx)) ; "cx"
       (printf "~a\n" (get-cell-value "E2" xlsx)))) ; 456.34
  
    ;; loop for row
    (with-input-from-xlsx-file
      "test2.xlsx"
      (lambda (xlsx)
        (load-sheet "Sheet1" xlsx)
        (with-row xlsx
                  (lambda (row)
                    (printf "~a\n" (first row))))))) ;; chenxiao 1
}  
            
@section{FAQ}

@subsection{Why simple-xlsx'write is so slow when I have much more rows?}

No, simple-xlsx's write is not slow(not so fast, perhaps).

If you find it turns slow when you have much rows, it perhaphs you prepare rows like below:

@verbatim {
  (let ([rows '()])
    (set! rows `(,@rows ,row))
 or (set! rows (append rows row)))
}

When rows's count is small, its ok, but when you have thousands rows, it's getting too slow to complete.

When work on big data, you should do like this:

@verbatim {
  (let loop ([rows '()]
             [line (read-line)])
                 
      (if (not (eof-object? line))
          (loop (cons (list (regexp-replace* #rx"\n|\r" line "")) rows) (read-line) (add1 count)))
          (send xlsx add-sheet (reverse rows) "Sheet1"))
}

Please notice you should use reverse list order when use cons to chain a list up.
           
