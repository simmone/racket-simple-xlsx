#lang racket

(require rackunit/text-ui)

(require rackunit "../../../main.rkt")

(define test-test7
  (test-suite
   "test-test7"

   (with-input-from-xlsx-file
    "test7.xlsx"
    (lambda (xlsx)
      (test-case
       "test-get-sheet-data"

       (load-sheet (car (get-sheet-names xlsx)) xlsx)
       (check-equal? (get-cell-value "D7" xlsx) "66260001")
      )))))

(run-tests test-test7)
