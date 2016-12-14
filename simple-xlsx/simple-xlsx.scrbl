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
    (send xlsx add-data-sheet #:sheet_name "Sheet1" #:sheet_data '(("chenxiao" "cx") (1 2)))
}

@subsubsection{set col width}

use set-data-sheet-col-width! method to set col's width

for example:
@verbatim{
  ;; set column A, B width: 50
  (send xlsx set-data-sheet-col-width! #:sheet_name "DataSheet" #:col_range "A-B" #:width 50)
}

@subsubsection{set cell background color}

use set-data-sheet-cell-color! method to set cell's background color

for example:
@verbatim{
  ;; set B2 to C3, 2X2, total 4 cells's color to FF0000(red)
  (send xlsx set-data-sheet-cell-color! #:sheet_name "DataSheet" #:cell_range "B2-C3" #:color "FF0000")
}

@subsection{Chart Sheet}

chart sheet is a sheet contains chart only.

chart sheet use data sheet's data to constuct chart.

chart type now can have: linechart, linechart3d, barchart, barchart3d, piechart, piechart3d

@subsubsection{add chart sheet}

default chart_type is linechart or set chart type

chart type is one of these: line, line3d, bar, bar3d, pie, pie3d

@verbatim{
  (send xlsx add-chart-sheet #:sheet_name "LineChart1" #:topic "Horizontal Data" #:x_topic "Kg")

  (send xlsx add-chart-sheet #:sheet_name "LineChart1" #:chart_type 'bar #:topic "Horizontal Data" #:x_topic "Kg")
}

@subsubsection{set-chart-x-data! and add-chart-serail!}

use this two methods to set chart's x axis data and y axis data

only one x axis data and multiple y axis data

@verbatim{
  (send xlsx set-chart-x-data! #:sheet_name "LineChart1" #:data_sheet_name "DataSheet" #:data_range "B1-D1")
  (send xlsx add-chart-serial! #:sheet_name "LineChart1" #:data_sheet_name "DataSheet" #:data_range "B2-D2" #:y_topic "CAT")
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

(require "../main.rkt")

(let ([xlsx (new xlsx%)])
  (send xlsx add-data-sheet 
        #:sheet_name "DataSheet" 
        #:sheet_data '(("month/brand" "201601" "201602" "201603")
                       ("CAT" 100 300 200)
                       ("Puma" 200 400 300)
                       ("Brooks" 300 500 400)
                       ))
  (send xlsx set-data-sheet-col-width! #:sheet_name "DataSheet" #:col_range "A-B" #:width 50)
  (send xlsx set-data-sheet-cell-color! #:sheet_name "DataSheet" #:cell_range "B2-C3" #:color "FF0000")
  (send xlsx set-data-sheet-cell-color! #:sheet_name "DataSheet" #:cell_range "C4-D4" #:color "0000FF")

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
     (printf "~a\n" (get-sheet-dimension xlsx)) ;(4 . 4)

     (printf "~a\n" (get-cell-value "A2" xlsx)) ;201601

     (printf "~a\n" (get-sheet-rows xlsx))))
     ; ((month/brand 201601 201602 201603) (CAT 100 300 200) (Puma 200 400 300) (Brooks 300 500 400))
  )

  (printf "~a\n" (sheet-name-rows "test.xlsx" "DataSheet"))
  ; ((month/brand 201601 201602 201603) (CAT 100 300 200) (Puma 200 400 300) (Brooks 300 500 400))

  (printf "~a\n" (sheet-ref-rows "test.xlsx" 0))
  ; ((month/brand 201601 201602 201603) (CAT 100 300 200) (Puma 200 400 300) (Brooks 300 500 400))
}
