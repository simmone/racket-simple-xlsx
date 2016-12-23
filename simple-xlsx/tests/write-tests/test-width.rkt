#lang racket

(require "../../main.rkt")

(let ([xlsx (new xlsx%)])
  (send xlsx add-data-sheet 
        #:sheet_name "DataSheet" 
        #:sheet_data '(("123456" "思考的3打开" "A思" "才开口快看看思考的福克斯" "思考的咖啡色的口sfskdfks" 1232 1232.23 1231231231 12312312.1231231)))

  (send xlsx set-data-sheet-col-width! #:sheet_name "DataSheet" #:col_range "A-B" #:width 20)

  (write-xlsx-file xlsx "test.xlsx")
  )
