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
      #:sheet_data '( ("number") (7.89) (3.21)))

(send xlsx add-data-sheet-col-style!
      #:sheet_name "s2"
      #:col_range "1"
      #:style '((numberPrecision . 2)))

(write-xlsx-file xlsx "bug.xlsx")
