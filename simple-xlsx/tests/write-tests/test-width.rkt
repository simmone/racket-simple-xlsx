#lang racket

(require "../../main.rkt")

(let ([xlsx (new xlsx%)])
  (send xlsx add-data-sheet 
        #:sheet_name "DataSheet" 
        #:sheet_data '(("123456" "abcdef" "ABCDEF" "才开口快看看思考的福克斯" "!'#$%&")))

  (send xlsx set-data-sheet-col-width! #:sheet_name "DataSheet" #:col_range "A" #:width 8)
  (send xlsx set-data-sheet-col-width! #:sheet_name "DataSheet" #:col_range "B" #:width 8)
  (send xlsx set-data-sheet-col-width! #:sheet_name "DataSheet" #:col_range "C" #:width 8)
  (send xlsx set-data-sheet-col-width! #:sheet_name "DataSheet" #:col_range "D" #:width 24)
  (send xlsx set-data-sheet-col-width! #:sheet_name "DataSheet" #:col_range "E" #:width 8)

  (write-xlsx-file xlsx "test.xlsx")
  )
