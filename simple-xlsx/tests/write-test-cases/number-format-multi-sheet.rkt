#lang racket

(require "../../main.rkt")

(define xlsx (new xlsx%))

(send xlsx add-data-sheet
      #:sheet_name "s1"
      #:sheet_data '( ("number") (1) (2)))

; BUG: numberPrecision on sheet "s2" does not work properly.
; Comment out this call, and numberPrecision on sheet "s2" will work.
(send xlsx add-data-sheet-col-style!
      #:sheet_name "s1"
      #:col_range "1"
      #:style '((numberPrecision . 0)))

(send xlsx add-data-sheet
      #:sheet_name "s2"
      #:sheet_data '( ("number") (7232.89) (-23423.21)))

(send xlsx add-data-sheet-cell-style!
      #:sheet_name "s2"
      #:cell_range "A1"
      #:style '((formatCode . "__@__@")))

(send xlsx add-data-sheet-cell-style!
      #:sheet_name "s2"
      #:cell_range "A2"
      #:style '((formatCode . "￥#,##0.00;[Red]￥-#,##0.00")))

(send xlsx add-data-sheet-cell-style!
      #:sheet_name "s2"
      #:cell_range "A3"
      #:style '((formatCode . "￥#,##0.00;[Red]￥-#,##0.00")))

(write-xlsx-file xlsx "template.xlsx")
