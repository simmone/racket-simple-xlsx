#lang racket

(require rackunit/text-ui)

(require rackunit "../../../main.rkt")

(require "../../../lib/lib.rkt")

(define test-test4
  (test-suite
   "test-test4"

   (with-input-from-xlsx-file
    "test4.xlsx"
    (lambda (xlsx)
      (test-case
       "test-get-sheet-data"

       (load-sheet "Sheet1" xlsx)
       (check-equal? (get-cell-value "A1" xlsx) "200008194477601")
       (check-equal? (get-cell-value "B1" xlsx) "20140425")
       (check-equal? (get-cell-value "C1" xlsx) "U0298")
       (check-equal? (get-cell-value "D1" xlsx) 100)
      )))))

(run-tests test-test4)