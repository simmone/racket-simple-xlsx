#lang racket

(require rackunit/text-ui)

(require rackunit)

(require "../../../main.rkt")

(require racket/runtime-path)
(define-runtime-path test_file "test10.xlsx")

(define test-test10
  (test-suite
   "test-test10"

   (test-case
    "test-get-cell-value"
    (with-input-from-xlsx-file
     test_file
     (lambda (xlsx)
       (load-sheet-ref 0 xlsx)
       (check-equal? (get-cell-value "A1" xlsx) 1)
       (check-equal? (get-cell-value "I1" xlsx) 9)
       )))

   (test-case
    "test-get-sheet-data"

    (check-equal? (sheet-ref-rows test_file 0) '((1 2 3 4 5 6 7 8 9)))
    )))

(run-tests test-test10)
