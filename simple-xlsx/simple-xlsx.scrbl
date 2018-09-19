#lang scribble/manual

@(require (for-label racket))

@title{Simple-Xlsx: Open Xml Spreadsheet(.xlsx) Reader and Writer}

@author+email["Chen Xiao" "chenxiao770117@gmail.com"]

simple-xlsx package is a package to read and write .xlsx format file.

.xlsx file is a open xml format file.

@table-of-contents[]

@section[#:tag "install"]{Install}

raco pkg install simple-xlsx

@section{Read}

read from a .xlsx file.

you can get a specific cell's value or loop for the whole sheet's rows.

@defmodule[simple-xlsx]
@(require (for-label simple-xlsx))

there is also a complete read and write example on github:@link["https://github.com/simmone/racket-simple-xlsx/blob/master/simple-xlsx/example/example.rkt"]{includedin the source}.

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
  load specified sheet by sheet name.
  must first called before other func, because any other func is based on specified sheet.
}

@defproc[(get-sheet-names
            [xlsx_handler (xlsx_handler)])
            list?]{
  get sheet names.
}

@defproc[(get-cell-value
            [cell_axis (string?)]
            [xlsx_handler (xlsx_handler)])
            any]{
  get cell value through cell's axis.
  cell axis: A1 B2 C3...
}

@defproc[(oa_date_number->date
            [oa_date_number (number?)])
            date?]{
  if knows cell's type is date, can use this function to convert to racket date? type.
  if not convert, xlsx's date type just is a number, like 43361.

  this function can convert number to a date? with precision to day only, the hour, minute and seconds set to 0.

  (oa_date_number->date 43359.1212121) to a date is 2018-9-16 00:00:00.
}

@defproc[(get-sheet-dimension
            [xlsx_handler (xlsx_handler)])
            pair?]{
  get current sheet's dimension, (cons row col)
  like (1 . 4)
}

@defproc[(get-sheet-rows
            [xlsx_handler (xlsx_handler)])
            list?]{
  get-sheet-rows get all rows from current loaded sheet
}

@defproc[(sheet-name-rows
            [xlsx_file_path (path-string?)]
            [sheet_name (string?)]
            )
            list?]{
  if, only if just want get a specific sheet name's data, no other operations on the xlsx file.

  this is the most simple func to get xlsx data.
}

@defproc[(sheet-ref-rows
            [xlsx_file_path (path-string?)]
            [sheet_index (exact-nonnegative-integer?)]
            )
            list?]{
  same as sheet-name-rows, use sheet index to specify sheet.
}

@section{Write}

write a xlsx file use xlsx% class.

use add-data-sheet method to add data type sheet to xlsx.

use add-chart-sheet method to add chart type sheet to xlsx.

@subsection{xlsx%}

xlsx% class represent a whole xlsx file's data.

it contains data sheet or chart sheet.

@subsection{Data Sheet}

data sheet is a sheet contains data only.

@subsubsection{add data sheet}

sheet data just a list contains list: (list (list cell ...) (list cell ...)...).

@verbatim{
  (let ([xlsx (new xlsx%)])
    (send xlsx add-data-sheet 
      #:sheet_name "Sheet1" 
      #:sheet_data '(("chenxiao" "cx") (1 2)))
}

@subsubsection{set col width}

column width is be set automatically by content's width.

if you want to set it manually, use set-data-sheet-col-width! method

for example:
@verbatim{
  ;; set column A, B width: 50
  (send xlsx set-data-sheet-col-width! 
    #:sheet_name "DataSheet" 
    #:col_range "A-B" #:width 50)
}

@subsection{Add Style to Data Sheet}

you can add various style to a data sheet.

includes background color, font style, number format, border style, date format.

use add-data-sheet-cell-style to add style to cells.

the last parameter style is pair list.

for example: @verbatim{'( (background . "FF0000") (fontSize . 20) )}

you can use add-data-sheet-cell-style multiple times, to a cell, it's a pile effect.

for example: 
@verbatim{
  (send xlsx add-data-sheet-cell-style! 
    #:sheet_name "DataSheet" 
    #:cell_range "B2-C3" 
    #:style '( (background . "FF0000") ))

  (send xlsx add-data-sheet-cell-style! 
    #:sheet_name "DataSheet" 
    #:cell_range "C3-D4" 
    #:style '( (fontSize . 30) ))
}

the C2's style is @verbatim{'( (background . "FF0000") )}

the D3's style is @verbatim{'( (fontSize . 30) )}

the C3's style is @verbatim{( (background . "FF0000") (fontSize . 30) )}

if set a cell the same style property multiple times, the last one works.

for example: 

@verbatim{
  (send xlsx add-data-sheet-cell-style! 
    #:sheet_name "DataSheet" 
    #:cell_range "B2-C3" 
    #:style '( (background . "FF0000") ))

  (send xlsx add-data-sheet-cell-style! 
    #:sheet_name "DataSheet" 
    #:cell_range "C3-D4" 
    #:style '( (background . "0000FF") ))
}
the C3's style is '( (background . "0000FF") ).

@subsubsection{backgroundColor}

rgb color or color name.

for example:
@verbatim{
  (send xlsx add-data-sheet-cell-style! 
    #:sheet_name "DataSheetWithStyle" 
    #:cell_range "A2-B3" 
    #:style '( (backgroundColor . "00C851") ))
}

@subsubsection{fontStyle}

fontSize: integer? default is 11.

fontColor: rgb color or colorname.

fontName: system font name.

for example:
@verbatim{
  (send xlsx add-data-sheet-cell-style! 
    #:sheet_name "DataSheetWithStyle" 
    #:cell_range "B3-C4" 
    #:style '( (fontSize . 20) (fontName . "Impact") (fontColor . "FF8800") ))
}

@subsubsection{numberFormat}

numberPrecision: non-exact-integer?

numberPercent: boolean?

numberThousands: boolean?

for example:
@verbatim{
  (send xlsx add-data-sheet-cell-style! 
    #:sheet_name "DataSheetWithStyle" 
    #:cell_range "E2-E2" 
    #:style '( 
              (numberPercent . #t) 
              (numberPrecision . 2) 
              (numberThousands . #t)))
}

@subsubsection{borderStyle}

borderDirection: @verbatim{'left 'right 'top 'bottom 'all}

boderStyle: 
@verbatim{
            'thin 'medium 'thick 'dashed 'thinDashed 

            'mediumDashed 'thickDashed 'double 'hair 'dotted 

            'dashDot 'dashDotDot 'mediumDashDot 'mediumDashDotDot 

            'slantDashDot
}

borderColor: rgb color or color name.

for example:
@verbatim{
  (send xlsx add-data-sheet-cell-style! 
    #:sheet_name "DataSheetWithStyle" 
    #:cell_range "B2-C4" 
    #:style '( (borderStyle . dashed) (borderColor . "blue")))
}

@subsubsection{dateFormat}

year: yyyy, month: mm, day: dd

for example:
@verbatim{
  (send xlsx add-data-sheet-cell-style! 
    #:sheet_name "DataSheetWithStyle" 
    #:cell_range "F2-F2" 
    #:style '( (dateFormat . "yyyy-mm-dd") ))

  (send xlsx add-data-sheet-cell-style! 
    #:sheet_name "DataSheetWithStyle" 
    #:cell_range "F2-F2" 
    #:style '( (dateFormat . "yyyy/mm/dd") ))
}

@subsection{Chart Sheet}

chart sheet is a sheet contains chart only.

chart sheet use data sheet's data to constuct chart.

chart type now can have: linechart, linechart3d, barchart, barchart3d, piechart, piechart3d

@subsubsection{add chart sheet}

default chart_type is linechart or set chart type

chart type is one of these: line, line3d, bar, bar3d, pie, pie3d

@verbatim{
  (send xlsx add-chart-sheet 
    #:sheet_name "LineChart1" 
    #:topic "Horizontal Data" 
    #:x_topic "Kg")

  (send xlsx add-chart-sheet 
    #:sheet_name "LineChart1" 
    #:chart_type 'bar 
    #:topic "Horizontal Data" 
    #:x_topic "Kg")
}

@subsubsection{set-chart-x-data! and add-chart-serail!}

use this two methods to set chart's x axis data and y axis data

only one x axis data and multiple y axis data

@verbatim{
  (send xlsx set-chart-x-data! 
    #:sheet_name "LineChart1" 
    #:data_sheet_name "DataSheet" 
    #:data_range "B1-D1")

  (send xlsx add-chart-serial! 
    #:sheet_name "LineChart1" 
    #:data_sheet_name "DataSheet" 
    #:data_range "B2-D2" #:y_topic "CAT")
}

@subsection{write file}

@defproc[(write-xlsx-file
            [xlsx (xlsx%)]
            [path (path-string?)])
            void?]{
  write xlsx% to xlsx file.
}

@section{Complete Example}

@verbatim{
#lang racket

(require simple-xlsx)

(require racket/date)

(let ([xlsx (new xlsx%)]
      [sheet_data 
        (list
          (list "month/brand" "201601" "201602" "201603" "201604" "201605")
          (list "CAT" 100 300 200 0.6934 
            (seconds->date (find-seconds 0 0 0 17 9 2018)))
          (list "Puma" 200 400 300 139999.89223 
            (seconds->date (find-seconds 0 0 0 18 9 2018)))
          (list "Asics" 300 500 400 23.34 
            (seconds->date (find-seconds 0 0 0 19 9 2018)))
          )])

  ;; add data          
  (send xlsx add-data-sheet 
    #:sheet_name "DataSheet" #:sheet_data sheet_data)

  ;; set column width manully
  (send xlsx set-data-sheet-col-width! 
    #:sheet_name "DataSheet" #:col_range "A-B" #:width 50)

  ;; add another data sheet
  (send xlsx add-data-sheet 
    #:sheet_name "DataSheetWithStyle" #:sheet_data sheet_data)
  (send xlsx set-data-sheet-col-width! 
    #:sheet_name "DataSheetWithStyle" #:col_range "A-B" #:width 50)

  ;; add various styles to data sheet
  (send xlsx add-data-sheet-cell-style! 
    #:sheet_name "DataSheetWithStyle" 
    #:cell_range "A2-B3" 
    #:style '( (backgroundColor . "00C851") ))
  (send xlsx add-data-sheet-cell-style! 
    #:sheet_name "DataSheetWithStyle" 
    #:cell_range "C3-D4" 
    #:style '( (backgroundColor . "AA66CC") ))

  (send xlsx add-data-sheet-cell-style! 
    #:sheet_name "DataSheetWithStyle" 
    #:cell_range "B3-C4" 
    #:style '( (fontSize . 20) (fontName . "Impact") ))
  (send xlsx add-data-sheet-cell-style! #:sheet_name "DataSheetWithStyle" 
    #:cell_range "B1-C3" 
    #:style '( (fontColor . "FF8800") ))

  (send xlsx add-data-sheet-cell-style! #:sheet_name "DataSheetWithStyle" 
    #:cell_range "E2-E2" 
    #:style '( (numberPercent . #t) ))
  (send xlsx add-data-sheet-cell-style! #:sheet_name "DataSheetWithStyle" 
    #:cell_range "E3-E3" 
    #:style '( (numberPrecision . 2) (numberThousands . #t) ))
  (send xlsx add-data-sheet-cell-style! #:sheet_name "DataSheetWithStyle" 
    #:cell_range "E4-E4" 
    #:style '( (numberPrecision . 0) ))

  (send xlsx add-data-sheet-cell-style! #:sheet_name "DataSheetWithStyle" 
    #:cell_range "B2-C4" 
    #:style '( (borderStyle . dashed) (borderColor . "blue") ))

  (send xlsx add-data-sheet-cell-style! #:sheet_name "DataSheetWithStyle" 
    #:cell_range "F2-F2" 
    #:style '( (dateFormat . "yyyy-mm-dd") ))
  (send xlsx add-data-sheet-cell-style! #:sheet_name "DataSheetWithStyle" 
    #:cell_range "F3-F3" 
    #:style '( (dateFormat . "yyyy/mm/dd") ))
  (send xlsx add-data-sheet-cell-style! #:sheet_name "DataSheetWithStyle" 
    #:cell_range "F4-F4" 
    #:style '( (dateFormat . "yyyy年mm月dd日") ))

  ;; LineChart
  (send xlsx add-chart-sheet 
    #:sheet_name "LineChart1" 
    #:topic "Horizontal Data" 
    #:x_topic "Kg")
  (send xlsx set-chart-x-data! 
    #:sheet_name "LineChart1" 
    #:data_sheet_name "DataSheet" 
    #:data_range "B1-D1")
  (send xlsx add-chart-serial! 
    #:sheet_name "LineChart1" 
    #:data_sheet_name "DataSheet" 
    #:data_range "B2-D2" 
    #:y_topic "CAT")
  (send xlsx add-chart-serial! 
    #:sheet_name "LineChart1" 
    #:data_sheet_name "DataSheet" 
    #:data_range "B3-D3" 
    #:y_topic "Puma")
  (send xlsx add-chart-serial! 
    #:sheet_name "LineChart1" 
    #:data_sheet_name "DataSheet" 
    #:data_range "B4-D4" 
    #:y_topic "Brooks")

  (send xlsx add-chart-sheet 
    #:sheet_name "LineChart2" 
    #:topic "Vertical Data" 
    #:x_topic "Kg")
  (send xlsx set-chart-x-data! 
    #:sheet_name "LineChart2" 
    #:data_sheet_name "DataSheet" 
    #:data_range "A2-A4" )
  (send xlsx add-chart-serial! 
    #:sheet_name "LineChart2" 
    #:data_sheet_name "DataSheet" 
    #:data_range "B2-B4" 
    #:y_topic "201601")
  (send xlsx add-chart-serial! 
    #:sheet_name "LineChart2" 
    #:data_sheet_name "DataSheet" 
    #:data_range "C2-C4" 
    #:y_topic "201602")
  (send xlsx add-chart-serial! 
    #:sheet_name "LineChart2" 
    #:data_sheet_name "DataSheet" 
    #:data_range "D2-D4" 
    #:y_topic "201603")

  (send xlsx add-chart-sheet 
    #:sheet_name "LineChart3D" 
    #:chart_type 'line3d 
    #:topic "LineChart3D" 
    #:x_topic "Kg")
  (send xlsx set-chart-x-data! 
    #:sheet_name "LineChart3D" 
    #:data_sheet_name "DataSheet" 
    #:data_range "A2-A4" )
  (send xlsx add-chart-serial! 
    #:sheet_name "LineChart3D" 
    #:data_sheet_name "DataSheet" 
    #:data_range "B2-B4" 
    #:y_topic "201601")
  (send xlsx add-chart-serial! 
    #:sheet_name "LineChart3D" 
    #:data_sheet_name "DataSheet" 
    #:data_range "C2-C4" 
    #:y_topic "201602")
  (send xlsx add-chart-serial! 
    #:sheet_name "LineChart3D" 
    #:data_sheet_name "DataSheet" 
    #:data_range "D2-D4" 
    #:y_topic "201603")

  ;; BarChart
  (send xlsx add-chart-sheet 
    #:sheet_name "BarChart" 
    #:chart_type 'bar 
    #:topic "BarChart" 
    #:x_topic "Kg")
  (send xlsx set-chart-x-data! 
    #:sheet_name "BarChart" 
    #:data_sheet_name "DataSheet" 
    #:data_range "B1-D1" )
  (send xlsx add-chart-serial! 
    #:sheet_name "BarChart" 
    #:data_sheet_name "DataSheet" 
    #:data_range "B2-D2" 
    #:y_topic "CAT")
  (send xlsx add-chart-serial! 
    #:sheet_name "BarChart" 
    #:data_sheet_name "DataSheet" 
    #:data_range "B3-D3" 
    #:y_topic "Puma")
  (send xlsx add-chart-serial! 
    #:sheet_name "BarChart" 
    #:data_sheet_name "DataSheet" 
    #:data_range "B4-D4" 
    #:y_topic "Brooks")

  ;; BarChart3D
  (send xlsx add-chart-sheet 
    #:sheet_name "BarChart3D" 
    #:chart_type 'bar3d 
    #:topic "BarChart3D" 
    #:x_topic "Kg")
  (send xlsx set-chart-x-data! 
    #:sheet_name "BarChart3D" 
    #:data_sheet_name "DataSheet" 
    #:data_range "B1-D1" )
  (send xlsx add-chart-serial! 
    #:sheet_name "BarChart3D" 
    #:data_sheet_name "DataSheet" 
    #:data_range "B2-D2" 
    #:y_topic "CAT")
  (send xlsx add-chart-serial! 
    #:sheet_name "BarChart3D" 
    #:data_sheet_name "DataSheet" 
    #:data_range "B3-D3" 
    #:y_topic "Puma")
  (send xlsx add-chart-serial! 
    #:sheet_name "BarChart3D" 
    #:data_sheet_name "DataSheet" 
    #:data_range "B4-D4" 
    #:y_topic "Brooks")

  ;; PieChart
  (send xlsx add-chart-sheet 
    #:sheet_name "PieChart" 
    #:chart_type 'pie 
    #:topic "PieChart" 
    #:x_topic "Kg")
  (send xlsx set-chart-x-data! 
    #:sheet_name "PieChart" 
    #:data_sheet_name "DataSheet" 
    #:data_range "B1-D1" )
  (send xlsx add-chart-serial! 
    #:sheet_name "PieChart" 
    #:data_sheet_name "DataSheet" 
    #:data_range "B2-D2" 
    #:y_topic "CAT")

  ;; PieChart3D
  (send xlsx add-chart-sheet 
    #:sheet_name "PieChart3D" 
    #:chart_type 'pie3d 
    #:topic "PieChart3D" 
    #:x_topic "Kg")
  (send xlsx set-chart-x-data! 
    #:sheet_name "PieChart3D" 
    #:data_sheet_name "DataSheet" 
    #:data_range "B1-D1" )
  (send xlsx add-chart-serial! 
    #:sheet_name "PieChart3D" 
    #:data_sheet_name "DataSheet" 
    #:data_range "B2-D2" 
    #:y_topic "CAT")

  (write-xlsx-file xlsx "test.xlsx")

  ;; read xlsx
  (with-input-from-xlsx-file
   "test.xlsx"
   (lambda (xlsx)
     (printf "~a\n" (get-sheet-names xlsx))
     ;("DataSheet" "LineChart1" "LineChart2" "LineChart3D" 
     ; "BarChart" "BarChart3D" "PieChart" "PieChart3D"))

     (load-sheet "DataSheet" xlsx)
     (printf "~a\n" (get-sheet-dimension xlsx)) ;(4 . 6)

     (printf "~a\n" (get-cell-value "A2" xlsx)) ;201601

     (let ([date_val (oa_date_number->date (get-cell-value "F2" xlsx))])
       (printf "~a,~a,~a\n" 
         (date-year date_val) 
         (date-month date_val) 
         (date-day date_val)))
     ; 2018,9,17

     (printf "~a\n" (get-sheet-rows xlsx))))
     ; ((month/brand 201601 201602 201603 201604 201605) 
     ;  (CAT 100 300 200 0.6934 43360) 
     ;  (Puma 200 400 300 139999.89223 43361) 
     ;  (Asics 300 500 400 23.34 43362))
  )

  ; result is same as (get-sheet-rows xlsx)
  (printf "~a\n" (sheet-name-rows "test.xlsx" "DataSheet"))

  (printf "~a\n" (sheet-ref-rows "test.xlsx" 0))
}
