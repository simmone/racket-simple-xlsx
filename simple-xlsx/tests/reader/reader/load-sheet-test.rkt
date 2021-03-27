#lang racket

(require rackunit/text-ui)

(require rackunit "../../../reader.rkt")

(require "../../../xlsx/xlsx.rkt")
(require "../../../sheet/sheet.rkt")

(require racket/runtime-path)
(define-runtime-path test1_file "test1.xlsx")

(define test-load-sheet
  (test-suite
   "test-load-sheet"
   
   (test-case
    "test-load-sheet"

    (with-input-from-xlsx-file
     test1_file
     (lambda (xlsx)
       (load-sheet "DataSheet" xlsx)
       
       (check-equal? (length (XLSX-sheet_list xlsx)) 1)
      )))

   (test-case
    "test-load-sheet-user-proc"

    (with-input-from-xlsx-file
     test1_file
     (lambda (xlsx)
       (load-sheet 
        "DataSheet" xlsx
        (lambda (sheet)
          (check-equal? (DATA-SHEET-dimension sheet) '(4 . 6))
          )))))
    ))

(run-tests test-load-sheet)
