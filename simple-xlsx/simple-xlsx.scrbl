#lang scribble/manual

@(require (for-label racket))
@(require (for-label simple-xlsx))

@(require scribble/example)

@(define example-eval (make-base-eval))
@(example-eval '(require simple-xlsx racket/date))

@title{Simple-Xlsx: Open Xml Spreadsheet(.xlsx) Reader and Writer}

@author+email["Chen Xiao" "chenxiao770117@gmail.com"]

@defmodule[simple-xlsx]

The @tt{simple-xlsx} package allows you to read and write spreadsheets in the @tt{.xlsx} file format
used by Microsoft Excel and LibreOffice. This an open XML file format.

@table-of-contents[]

@section[#:tag "install"]{Install}

@codeblock{raco pkg install simple-xlsx}

@section{Read}

Functions for reading from a @filepath{.xlsx} file.

You can get a specific cell's value or loop for the whole sheet's rows.

There is also a complete read and write example 
@link["https://github.com/simmone/racket-simple-xlsx/blob/master/simple-xlsx/example/example.rkt"]{included
in the GitHub source}.

@defproc[(with-input-from-xlsx-file
              [xlsx_file_path path-string?]
              [user-proc (-> (is-a?/c xlsx%) void?)]
              )
            void?]{
  Loads a @filepath{.xlsx} file and calls @racket[_user-proc] with the @racket[xlsx%] object as its only
  argument.
}

@defproc[(load-sheet
           [sheet_name string?]
           [xlsx_handler (is-a?/c xlsx%)]
           )
           void?]{
  Load a sheet specified by its sheet name.
  
  This must be called before attempting to read any cell values.
}

@defproc[(get-sheet-names
            [xlsx_handler (is-a?/c xlsx%)]
            )
            (listof string?)]{
  Returns a list of sheet names.
}

@defproc[(get-cell-value
            [cell_axis string?]
            [xlsx_handler (is-a?/c xlsx%)]
            )
            any]{
  Returns the value of a specific cell. The @racket[_cell-axis] should be in the “A1” reference
  style.

  Example:

  @racketblock[
  (with-input-from-xlsx-file "workbook.xlsx"
    (lambda (xlsx)
      (load-sheet "Sheet1" xlsx)
      (get-cell-value "C12" xlsx)))
  ]
}

@defproc[(get-cell-formula
            [cell_axis string?]
            [xlsx_handler (is-a?/c xlsx%)]
            )
            string?]{
  Get a cell's formula (as opposed to the calculated value of the formula). If the cell has no formula, this will return an empty string.

  The @racket[_cell-axis] should be in the “A1” reference style.
  
  Limitations: Currently does not support array or shared formulae.
}

@defproc[(oa_date_number->date
            [oa_date_number number?]
            )
            date?]{
Convert an @tt{xlsx} cell's "date" value into Racket's @racket[date?] struct. Any fractional portion of @racket[_oa_date_number] is ignored; this function's precision is to the day only.
                                            
Cells with a "date" type in @tt{xlsx} files are a plain number representing the count of days since 0 January 1900.

@examples[#:eval example-eval
  (date->string (oa_date_number->date 43359.1212121))
]
}

@defproc[(get-sheet-dimension
            [xlsx_handler (is-a?/c xlsx%)]
            )
            pair?]{
  Returns the current sheet's dimension as @racket[(cons _row _col)], such as @racket['(1 . 4)].
}

@defproc[(get-sheet-rows
            [xlsx_handler (is-a?/c xlsx%)]
            )
            list?]{
  get-sheet-rows get all rows from current loaded sheet
}

@defproc[(sheet-name-rows
            [xlsx_file_path path-string?]
            [sheet_name string?]
            )
            list?]{
  if, only if just want get a specific sheet name's data, no other operations on the xlsx file.

  this is the most simple func to get xlsx data.
}

@defproc[(sheet-ref-rows
            [xlsx_file_path path-string?]
            [sheet_index exact-nonnegative-integer?]
            )
            list?]{
  same as sheet-name-rows, use sheet index to specify sheet.
}

@section{Write}

write a xlsx file use xlsx% class.

use add-data-sheet method to add data type sheet to xlsx.

use add-chart-sheet method to add chart type sheet to xlsx.

@subsection{xlsx%}

@defclass[xlsx% object% ()]{

The @racket[xlsx%] class represents a whole xlsx file's data.

It contains data sheet or chart sheet.


@defmethod[(add-data-sheet [#:sheet_name sheet string?]
                           [#:sheet_data cells (listof (listof any/c))]) void?]{

Adds a data sheet (as opposed to a chart sheet) which holds normal data in cells.

Example:

@codeblock{
  (let ([xlsx (new xlsx%)])
    (send xlsx add-data-sheet 
      #:sheet_name "Sheet1" 
      #:sheet_data '(("chenxiao" "cx") (1 2))))
}

}
                                                                         
@defmethod[(set-data-sheet-col-width! [#:sheet_name sheet string?]
                                      [#:col_range cols string?]
                                      [#:width width number?]) void?]{

Manually set the width of one or more columns.

Note that by default, column widths are set automatically by their content. Use this method to
override the automatic sizing.

Example:

@codeblock{
  ;; set column A, B width: 50
  (send xlsx set-data-sheet-col-width! 
    #:sheet_name "DataSheet" 
    #:col_range "A-B" #:width 50)
}

}

@defmethod[(set-data-sheet-row-height! [#:sheet_name sheet string?]
                                       [#:row_range rows string?]
                                       [#:height height number?]) void?]{

Set the height of specified rows.

Example:

@codeblock{
  (send xlsx set-data-sheet-row-height!
    #:sheet_name "DataSheetWithStyle2"
    #:row_range "2-4" #:height 30)
}

}

@defmethod[(set-data-sheet-freeze-pane! [#:sheet_name sheet string?]
                                        [#:range range (cons/c exact-nonnegative-integer?
                                                               exact-nonnegative-integer?)]) void?]{

“Freezes” the given number of rows (counting from the top) and columns (counting from the left).

Example:

@codeblock{
  ;; freeze 1 row and 1 col
  (send xlsx set-data-sheet-freeze-pane! #:sheet_name "DataSheet" #:range '(1 . 1))
}

}



@defmethod[(add-data-sheet-cell-style! [#:sheet_name sheet string?]
                                       [#:cell_range range string?]
                                       [#:style style (listof (cons symbol? any/c))]) void?]{

add-data-sheet-row-style! set rows styles.

add-data-sheet-col-style! set cols styles.

styles format: @verbatim{'( (backgroundColor . "FF0000") (fontSize . 20) )}

you can set cell, row, col style any times, it's a pile effect.

it means:

if the latter style has same style property, it'll overwrite this property.

if not, it'll add this property.

it also means the order you set style is important.

@codeblock{
  (send xlsx add-data-sheet-cell-style! 
    #:sheet_name "DataSheet" 
    #:cell_range "B2-C3" 
    #:style '( (backgroundColor . "FF0000") ))

  (send xlsx add-data-sheet-cell-style! 
    #:sheet_name "DataSheet" 
    #:cell_range "C3-D4" 
    #:style '( (fontSize . 30) ))

  (send xlsx add-data-sheet-row-style!
    #:sheet_name "DataSheetWithStyle2"
    #:row_range "1-3" #:style '( (backgroundColor . "00C851") ))

  (send xlsx add-data-sheet-col-style!
    #:sheet_name "DataSheetWithStyle2"
    #:col_range "4-6" #:style '( (backgroundColor . "AA66CC") ))
}

the C2's style is @verbatim{'( (backgroundColor . "AA66CC") )}

the D3's style is @verbatim{'( (backgroundColor . "AA66CC") (fontSize . 30) )}

the C3's style is @verbatim{( (backgroundColor . "00C851") (fontSize . 30) )}

@codeblock{
  (send xlsx add-data-sheet-cell-style! 
    #:sheet_name "DataSheet" 
    #:cell_range "B2-C3" 
    #:style '( (backgroundColor . "FF0000") ))

  (send xlsx add-data-sheet-cell-style! 
    #:sheet_name "DataSheet" 
    #:cell_range "C3-D4" 
    #:style '( (backgroundColor . "0000FF") ))
}

the C3's style is '( (backgroundColor . "0000FF") )
}

} @; defclass
                                                                                            
@subsubsection{Background Color}

@verbatim{'backgroundColor}

rgb color or color name.

@codeblock{
  (send xlsx add-data-sheet-cell-style! 
    #:sheet_name "DataSheetWithStyle" 
    #:cell_range "A2-B3" 
    #:style '( (backgroundColor . "00C851") ))
}

@subsubsection{Font Style}

@verbatim{'fontSize 'fontColor 'fontName}

fontSize: integer? default is 11.

fontColor: rgb color or colorname.

fontName: system font name.

@codeblock{
  (send xlsx add-data-sheet-cell-style! 
    #:sheet_name "DataSheetWithStyle" 
    #:cell_range "B3-C4" 
    #:style '( (fontSize . 20) (fontName . "Impact") (fontColor . "FF8800") ))
}

@subsubsection{Number Format}

@verbatim{'numberPercent 'numberPrecision 'numberThousands}

numberPrecision: non-exact-integer?

numberPercent: boolean?

numberThousands: boolean?

@codeblock{
  (send xlsx add-data-sheet-cell-style! 
    #:sheet_name "DataSheetWithStyle" 
    #:cell_range "E2-E2" 
    #:style '( 
              (numberPercent . #t) 
              (numberPrecision . 2) 
              (numberThousands . #t)
              ))
}

@subsubsection{Border Style}

@verbatim{'borderStyle 'borderColor}

borderDirection: @verbatim{'left 'right 'top 'bottom 'all}

boderStyle: 
@verbatim{
            'thin 'medium 'thick 'dashed 'thinDashed 

            'mediumDashed 'thickDashed 'double 'hair 'dotted 

            'dashDot 'dashDotDot 'mediumDashDot 'mediumDashDotDot 

            'slantDashDot
}

borderColor: rgb color or color name.

@codeblock{
  (send xlsx add-data-sheet-cell-style! 
    #:sheet_name "DataSheetWithStyle" 
    #:cell_range "B2-C4" 
    #:style '( (borderStyle . dashed) (borderColor . "blue")))
}

@subsubsection{Date Format}

@verbatim{'dateFormat}

year: yyyy, month: mm, day: dd

@codeblock{
  (send xlsx add-data-sheet-cell-style! 
    #:sheet_name "DataSheetWithStyle" 
    #:cell_range "F2-F2" 
    #:style '( (dateFormat . "yyyy-mm-dd") ))

  (send xlsx add-data-sheet-cell-style! 
    #:sheet_name "DataSheetWithStyle" 
    #:cell_range "F2-F2" 
    #:style '( (dateFormat . "yyyy/mm/dd") ))
}

@subsubsection{Cell Alignment}

@verbatim{'horizontalAlign 'verticalAlign}

horizontalAlign: 'left 'right 'center

verticalAlign: 'top 'bottom 'middle

@codeblock{
  (send xlsx add-data-sheet-cell-style!
    #:sheet_name "DataSheetWithStyle2"
    #:cell_range "G5"
    #:style '( (horizontalAlign . left) ))
}

}
@subsection{Chart Sheet}

chart sheet is a sheet contains chart only.

chart sheet use data sheet's data to constuct chart.

chart type now can have: linechart, linechart3d, barchart, barchart3d, piechart, piechart3d

@subsubsection{add chart sheet}

default chart_type is linechart or set chart type

chart type is one of these: line, line3d, bar, bar3d, pie, pie3d

@codeblock{
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

set-chart-x-data! and add-chart-serail!:

use this two methods to set chart's x axis data and y axis data

only one x axis data and multiple y axis data

@codeblock{
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
            [xlsx (is-a?/c xlsx%)]
            [path path-string?])
            void?]{
  write xlsx% to xlsx file.
}

@section{From Read to Write}

@defproc[(from-read-to-write-xlsx
            [read_xlsx (is-a?/c xlsx%)])
            (is-a?/c xlsx%)]{
  convert read xlsx object to write xlsx object.
}

@codeblock{
  (with-input-from-xlsx-file
   "test.xlsx"
   (lambda (xlsx)
     (let ([write_xlsx (from-read-to-write-xlsx xlsx)])
       (send write_xlsx set-data-sheet-col-width!
             #:sheet_name "DataSheet"
             #:col_range "A-F" #:width 20)
       (write-xlsx-file write_xlsx "write_back.xlsx"))))
}

@section{Complete Example}

@codeblock{
#lang racket

(require simple-xlsx)

(require racket/date)

(let ([xlsx (new xlsx%)]
      [sheet_data (list
                   (list "month/brand" "201601" "201602" "201603" "201604" "201605")
                   (list "CAT" 100 300 200 0.6934 (seconds->date (find-seconds 0 0 0 17 9 2018)))
                   (list "Puma" 200 400 300 139999.89223 (seconds->date (find-seconds 0 0 0 18 9 2018)))
                   (list "Asics" 300 500 400 23.34 (seconds->date (find-seconds 0 0 0 19 9 2018))))]
      [sheet_data2 (list
                    (list "month/brand" "201601" "201602" "201603" "201604" "201605" "")
                    (list "CAT" 100 300 200 0.6934 (seconds->date (find-seconds 0 0 0 17 9 2018)) "")
                    (list "Puma" 200 400 300 139999.89223 (seconds->date (find-seconds 0 0 0 18 9 2018)) "")
                    (list "Asics" 300 500 400 23.34 (seconds->date (find-seconds 0 0 0 19 9 2018)) "")
                    (list "" "" "" "" "" "" "Left")
                    (list "" "" "" "" "" "" "Right")
                    (list "" "" "" "" "" "" "Center")
                    (list "" "" "" "" "" "" "Top")
                    (list "" "" "" "" "" "" "Bottom")
                    (list "" "" "" "" "" "" "Middle")
                    (list "" "" "" "" "" "" "Center/Middle")
                    )])

  (send xlsx add-data-sheet #:sheet_name "DataSheet" #:sheet_data sheet_data)
  (send xlsx set-data-sheet-col-width! #:sheet_name "DataSheet" #:col_range "A-B" #:width 50)
  (send xlsx set-data-sheet-freeze-pane! #:sheet_name "DataSheet" #:range '(1 . 1))

  (send xlsx add-data-sheet #:sheet_name "DataSheetWithStyle" #:sheet_data sheet_data)
  (send xlsx set-data-sheet-col-width! #:sheet_name "DataSheetWithStyle" #:col_range "A-B" #:width 50)
  (send xlsx set-data-sheet-row-height! #:sheet_name "DataSheetWithStyle" #:row_range "3-4" #:height 30)
  (send xlsx set-data-sheet-col-width! #:sheet_name "DataSheetWithStyle" #:col_range "F" #:width 20)
  (send xlsx add-data-sheet-cell-style! #:sheet_name "DataSheetWithStyle" #:cell_range "A2-B3" #:style '( (backgroundColor . "00C851") ))
  (send xlsx add-data-sheet-cell-style! #:sheet_name "DataSheetWithStyle" #:cell_range "C3-D4" #:style '( (backgroundColor . "AA66CC") ))
  (send xlsx add-data-sheet-cell-style! #:sheet_name "DataSheetWithStyle" #:cell_range "B3-C4" #:style '( (fontSize . 20) (fontName . "Impact")))
  (send xlsx add-data-sheet-cell-style! #:sheet_name "DataSheetWithStyle" #:cell_range "B1-C3" #:style '( (fontColor . "FF8800") ))
  (send xlsx add-data-sheet-cell-style! #:sheet_name "DataSheetWithStyle" #:cell_range "E2-E2" #:style '( (numberPercent . #t) ))
  (send xlsx add-data-sheet-cell-style! #:sheet_name "DataSheetWithStyle" #:cell_range "E3-E3" #:style '( (numberPrecision . 2) (numberThousands . #t)))
  (send xlsx add-data-sheet-cell-style! #:sheet_name "DataSheetWithStyle" #:cell_range "E4-E4" #:style '( (numberPrecision . 0) ))
  (send xlsx add-data-sheet-cell-style! #:sheet_name "DataSheetWithStyle" #:cell_range "B2-C4" #:style '( (borderStyle . dashed) (borderColor . "blue")))
  (send xlsx add-data-sheet-cell-style! #:sheet_name "DataSheetWithStyle" #:cell_range "F2-F2" #:style '( (dateFormat . "yyyy-mm-dd") ))
  (send xlsx add-data-sheet-cell-style! #:sheet_name "DataSheetWithStyle" #:cell_range "F3-F3" #:style '( (dateFormat . "yyyy/mm/dd") ))
  (send xlsx add-data-sheet-cell-style! #:sheet_name "DataSheetWithStyle" #:cell_range "F4-F4" #:style '( (dateFormat . "yyyy年mm月dd日") ))

  (send xlsx add-data-sheet #:sheet_name "DataSheetWithStyle2" #:sheet_data sheet_data2)
  (send xlsx set-data-sheet-col-width! #:sheet_name "DataSheetWithStyle2" #:col_range "1-1" #:width 20)
  (send xlsx set-data-sheet-row-height! #:sheet_name "DataSheetWithStyle2" #:row_range "2-4" #:height 30)
  (send xlsx set-data-sheet-col-width! #:sheet_name "DataSheetWithStyle2" #:col_range "2-6" #:width 10)
  (send xlsx add-data-sheet-row-style! #:sheet_name "DataSheetWithStyle2" #:row_range "1-3" #:style '( (backgroundColor . "00C851") ))
  (send xlsx add-data-sheet-col-style! #:sheet_name "DataSheetWithStyle2" #:col_range "1-6" #:style '( (backgroundColor . "AA66CC") ))
  (send xlsx add-data-sheet-cell-style! #:sheet_name "DataSheetWithStyle2" #:cell_range "B1-C3" #:style '( (backgroundColor . "FF8800") ))

  (send xlsx set-data-sheet-col-width! #:sheet_name "DataSheetWithStyle2" #:col_range "7" #:width 50)
  (send xlsx set-data-sheet-row-height! #:sheet_name "DataSheetWithStyle2" #:row_range "5-11" #:height 50)
  (send xlsx add-data-sheet-cell-style! #:sheet_name "DataSheetWithStyle2" #:cell_range "G5" #:style '( (horizontalAlign . left) ))
  (send xlsx add-data-sheet-cell-style! #:sheet_name "DataSheetWithStyle2" #:cell_range "G6" #:style '( (horizontalAlign . right) ))
  (send xlsx add-data-sheet-cell-style! #:sheet_name "DataSheetWithStyle2" #:cell_range "G7" #:style '( (horizontalAlign . center) ))
  (send xlsx add-data-sheet-cell-style! #:sheet_name "DataSheetWithStyle2" #:cell_range "G8" #:style '( (verticalAlign . top) ))
  (send xlsx add-data-sheet-cell-style! #:sheet_name "DataSheetWithStyle2" #:cell_range "G9" #:style '( (verticalAlign . bottom) ))
  (send xlsx add-data-sheet-cell-style! #:sheet_name "DataSheetWithStyle2" #:cell_range "G10" #:style '( (verticalAlign . middle) ))
  (send xlsx add-data-sheet-cell-style! #:sheet_name "DataSheetWithStyle2" #:cell_range "G11"
        #:style '( (horizontalAlign . center) (verticalAlign . middle) ))

  (send xlsx add-chart-sheet #:sheet_name "LineChart1" #:topic "Horizontal Data" #:x_topic "Kg")
  (send xlsx set-chart-x-data! #:sheet_name "LineChart1" #:data_sheet_name "DataSheet" #:data_range "B1-D1")
  (send xlsx add-chart-serial! #:sheet_name "LineChart1" #:data_sheet_name "DataSheet" #:data_range "B2-D2" #:y_topic "CAT")
  (send xlsx add-chart-serial! #:sheet_name "LineChart1" #:data_sheet_name "DataSheet" #:data_range "B3-D3" #:y_topic "Puma")
  (send xlsx add-chart-serial! #:sheet_name "LineChart1" #:data_sheet_name "DataSheet" #:data_range "B4-D4" #:y_topic "Brooks")

  (send xlsx add-chart-sheet #:sheet_name "LineChart2" #:topic "Vertical Data" #:x_topic "Kg")
  (send xlsx set-chart-x-data! #:sheet_name "LineChart2" #:data_sheet_name "DataSheet" #:data_range "A2-A4" )
  (send xlsx add-chart-serial! #:sheet_name "LineChart2" #:data_sheet_name "DataSheet" #:data_range "B2-B4" #:y_topic "201601")
  (send xlsx add-chart-serial! #:sheet_name "LineChart2" #:data_sheet_name "DataSheet" #:data_range "C2-C4" #:y_topic "201602")
  (send xlsx add-chart-serial! #:sheet_name "LineChart2" #:data_sheet_name "DataSheet" #:data_range "D2-D4" #:y_topic "201603")

  (send xlsx add-chart-sheet #:sheet_name "LineChart3D" #:chart_type 'line3d #:topic "LineChart3D" #:x_topic "Kg")
  (send xlsx set-chart-x-data! #:sheet_name "LineChart3D" #:data_sheet_name "DataSheet" #:data_range "A2-A4" )
  (send xlsx add-chart-serial! #:sheet_name "LineChart3D" #:data_sheet_name "DataSheet" #:data_range "B2-B4" #:y_topic "201601")
  (send xlsx add-chart-serial! #:sheet_name "LineChart3D" #:data_sheet_name "DataSheet" #:data_range "C2-C4" #:y_topic "201602")
  (send xlsx add-chart-serial! #:sheet_name "LineChart3D" #:data_sheet_name "DataSheet" #:data_range "D2-D4" #:y_topic "201603")

  (send xlsx add-chart-sheet #:sheet_name "BarChart" #:chart_type 'bar #:topic "BarChart" #:x_topic "Kg")
  (send xlsx set-chart-x-data! #:sheet_name "BarChart" #:data_sheet_name "DataSheet" #:data_range "B1-D1" )
  (send xlsx add-chart-serial! #:sheet_name "BarChart" #:data_sheet_name "DataSheet" #:data_range "B2-D2" #:y_topic "CAT")
  (send xlsx add-chart-serial! #:sheet_name "BarChart" #:data_sheet_name "DataSheet" #:data_range "B3-D3" #:y_topic "Puma")
  (send xlsx add-chart-serial! #:sheet_name "BarChart" #:data_sheet_name "DataSheet" #:data_range "B4-D4" #:y_topic "Brooks")

  (send xlsx add-chart-sheet #:sheet_name "BarChart3D" #:chart_type 'bar3d #:topic "BarChart3D" #:x_topic "Kg")
  (send xlsx set-chart-x-data! #:sheet_name "BarChart3D" #:data_sheet_name "DataSheet" #:data_range "B1-D1" )
  (send xlsx add-chart-serial! #:sheet_name "BarChart3D" #:data_sheet_name "DataSheet" #:data_range "B2-D2" #:y_topic "CAT")
  (send xlsx add-chart-serial! #:sheet_name "BarChart3D" #:data_sheet_name "DataSheet" #:data_range "B3-D3" #:y_topic "Puma")
  (send xlsx add-chart-serial! #:sheet_name "BarChart3D" #:data_sheet_name "DataSheet" #:data_range "B4-D4" #:y_topic "Brooks")

  (send xlsx add-chart-sheet #:sheet_name "PieChart" #:chart_type 'pie #:topic "PieChart" #:x_topic "Kg")
  (send xlsx set-chart-x-data! #:sheet_name "PieChart" #:data_sheet_name "DataSheet" #:data_range "B1-D1" )
  (send xlsx add-chart-serial! #:sheet_name "PieChart" #:data_sheet_name "DataSheet" #:data_range "B2-D2" #:y_topic "CAT")

  (send xlsx add-chart-sheet #:sheet_name "PieChart3D" #:chart_type 'pie3d #:topic "PieChart3D" #:x_topic "Kg")
  (send xlsx set-chart-x-data! #:sheet_name "PieChart3D" #:data_sheet_name "DataSheet" #:data_range "B1-D1" )
  (send xlsx add-chart-serial! #:sheet_name "PieChart3D" #:data_sheet_name "DataSheet" #:data_range "B2-D2" #:y_topic "CAT")

  (write-xlsx-file xlsx "test.xlsx")

  (with-input-from-xlsx-file
   "test.xlsx"
   (lambda (xlsx)
     (printf "~a\n" (get-sheet-names xlsx))
     ;("DataSheet" "LineChart1" "LineChart2" "LineChart3D" "BarChart" "BarChart3D" "PieChart" "PieChart3D"))

     (load-sheet "DataSheet" xlsx)
     (printf "~a\n" (get-sheet-dimension xlsx)) ;(4 . 6)

     (printf "~a\n" (get-cell-value "A2" xlsx)) ;201601

     (let ([date_val (oa_date_number->date (get-cell-value "F2" xlsx))])
       (printf "~a,~a,~a\n" (date-year date_val) (date-month date_val) (date-day date_val)))
     ; 2018,9,17

     (printf "~a\n" (get-sheet-rows xlsx))
     ; ((month/brand 201601 201602 201603 201604 201605) (CAT 100 300 200 0.6934 43360) (Puma 200 400 300 139999.89223 43361) (Asics 300 500 400 23.34 43362))
     
     ))
  )

  (with-input-from-xlsx-file
   "test.xlsx"
   (lambda (xlsx)
     (let ([write_xlsx (from-read-to-write-xlsx xlsx)])
       (send write_xlsx set-data-sheet-col-width!
             #:sheet_name "DataSheet"
             #:col_range "A-F" #:width 20)
       (write-xlsx-file write_xlsx "write_back.xlsx"))))

  (printf "~a\n" (sheet-name-rows "test.xlsx" "DataSheet"))
  ; ((month/brand 201601 201602 201603 201604 201605) (CAT 100 300 200 0.6934 43360) (Puma 200 400 300 139999.89223 43361) (Asics 300 500 400 23.34 43362))

  (printf "~a\n" (sheet-ref-rows "test.xlsx" 0))
  ; ((month/brand 201601 201602 201603 201604 201605) (CAT 100 300 200 0.6934 43360) (Puma 200 400 300 139999.89223 43361) (Asics 300 500 400 23.34 43362))
}
