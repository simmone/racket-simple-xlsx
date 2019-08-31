#lang racket

;; (require simple-xlsx racket/date)

(require racket/date)

(require "../../main.rkt")

(let ([xlsx (new xlsx%)]
      [sheet_data
       (list
        (list "month/brand" "201601" "201602" "201603" "201604" "201605")
        (list "CAT" 100 300 200 0.6934 (seconds->date (find-seconds 0 0 0 17 9 2018)))
        (list "Puma" 200 400 300 139999.89223 (seconds->date (find-seconds 0 0 0 18 9 2018)))
        (list "Asics" 300 500 400 23.34 (seconds->date (find-seconds 0 0 0 19 9 2018)))
        )])

  ;; add data
  (send xlsx add-data-sheet #:sheet_name "DataSheet" #:sheet_data sheet_data)

  ;; (Part 1) Set date formats — no color info...
  (send xlsx add-data-sheet-cell-style! #:sheet_name "DataSheet" #:cell_range "F2-F2" #:style '( (dateFormat . "yyyy-mm-dd") ))
  (send xlsx add-data-sheet-cell-style! #:sheet_name "DataSheet" #:cell_range "F3-F3" #:style '( (dateFormat . "yyyy/mm/dd") ))
  (send xlsx add-data-sheet-cell-style! #:sheet_name "DataSheet" #:cell_range "F4-F4" #:style '( (dateFormat . "yyyy年mm月dd日") ))

  ;; cell -> row -> col
  (send xlsx add-data-sheet-cell-style! #:sheet_name "DataSheet" #:cell_range "F4-F4" #:style '( (backgroundColor . "red")))
  (send xlsx add-data-sheet-row-style! #:sheet_name "DataSheet" #:row_range "1-3" #:style '( (backgroundColor . "FEFFC1") ))
  (send xlsx add-data-sheet-col-style! #:sheet_name "DataSheet" #:col_range "2-6" #:style '( (backgroundColor . "00CCFF") ))
  (write-xlsx-file xlsx "cell-row-col.xlsx")

  ;; cell -> col -> row
  (send xlsx add-data-sheet-cell-style! #:sheet_name "DataSheet" #:cell_range "F4-F4" #:style '( (backgroundColor . "red")))
  (send xlsx add-data-sheet-col-style! #:sheet_name "DataSheet" #:col_range "2-6" #:style '( (backgroundColor . "00CCFF") ))
  (send xlsx add-data-sheet-row-style! #:sheet_name "DataSheet" #:row_range "1-3" #:style '( (backgroundColor . "FEFFC1") ))
  (write-xlsx-file xlsx "cell-col-row.xlsx")

  ;; col -> row -> cell
  (send xlsx add-data-sheet-col-style! #:sheet_name "DataSheet" #:col_range "2-6" #:style '( (backgroundColor . "00CCFF") ))
  (send xlsx add-data-sheet-row-style! #:sheet_name "DataSheet" #:row_range "1-3" #:style '( (backgroundColor . "FEFFC1") ))
  (send xlsx add-data-sheet-cell-style! #:sheet_name "DataSheet" #:cell_range "F4-F4" #:style '( (backgroundColor . "FF0000")))
  (write-xlsx-file xlsx "col-row-cell.xlsx")

  ;; row -> col -> cell
  (send xlsx add-data-sheet-row-style! #:sheet_name "DataSheet" #:row_range "1-3" #:style '( (backgroundColor . "FEFFC1") ))
  (send xlsx add-data-sheet-col-style! #:sheet_name "DataSheet" #:col_range "2-6" #:style '( (backgroundColor . "00CCFF") ))
  (send xlsx add-data-sheet-cell-style! #:sheet_name "DataSheet" #:cell_range "F4-F4" #:style '( (backgroundColor . "FF0000")))
  (write-xlsx-file xlsx "row-col-cell.xlsx")

  )

