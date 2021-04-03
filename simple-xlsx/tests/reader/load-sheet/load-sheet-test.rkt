#lang racket

(require rackunit/text-ui)

(require rackunit "../../../reader/load-sheet.rkt")

(require "../../../sheet/sheet.rkt")

(require racket/runtime-path)
(define-runtime-path sheet_file "sheet1.xml")

(define test-load-sheet-file
  (test-suite
   "test-load-sheet-file"
   
   (test-case
    "test-load-data-sheet-file"

    (let ([sheet (load-data-sheet-file sheet_file)])

      (check-equal? (DATA-SHEET-dimension sheet) '(4 . 6))
      (check-equal? (hash-count (DATA-SHEET-rvtsf_map sheet)) 24)

      ))))

(run-tests test-load-sheet-file)
