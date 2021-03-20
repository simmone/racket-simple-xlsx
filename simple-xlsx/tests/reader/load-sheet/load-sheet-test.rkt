#lang racket

(require rackunit/text-ui)

(require rackunit "../../../reader/load-sheet.rkt")

(require racket/runtime-path)
(define-runtime-path sheet_file "sheet1.xml")

(define test-load-sheet-file
  (test-suite
   "test-load-sheet-file"
   
   (test-case
    "test-load-data-sheet-file"

    (let-values ([sheet (load-data-sheet-file sheet_file)])

      (check-equal? (length (DATA-SHEET-rows sheet)) 4)
      (check-equal? (length (car (DATA-SHEET-rows sheet))) 4)
      (check-equal? (hash-count (DATA-SHEET-data_map sheet)) 16)

      ))))

(run-tests test-load-sheet-file)
