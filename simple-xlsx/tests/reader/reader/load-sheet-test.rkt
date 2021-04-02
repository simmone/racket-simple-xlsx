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
     (lambda ()
       (load-sheet "DataSheet")
       
       (check-equal? (length (XLSX-sheet_list (*CURRENT_XLSX*))) 1)
      )))

   (test-case
    "test-load-sheet-user-proc"

    (with-input-from-xlsx-file
     test1_file
     (lambda ()
       (load-sheet 
        "DataSheet"
        (lambda ()
          (check-equal? (sheet-dimension) '(4 . 6))

          (check-equal? (get-cell-value "A1") "month/brand")
          (check-equal? (get-cell-value "a1") "month/brand")
          (check-equal? (get-cell-value "a") "")
          )))))
    ))

(run-tests test-load-sheet)
