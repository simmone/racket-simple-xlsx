#lang racket

(require rackunit/text-ui)

(require rackunit "../../../main.rkt")

(require racket/runtime-path)
(define-runtime-path test_file "test5.xlsx")

(define test-test5
  (test-suite
   "test-test5"

   (with-input-from-xlsx-file
    test_file
    (lambda (xlsx)
      (test-case
       "test-get-sheet-data"

       (load-sheet "Sheet1" xlsx)
       (check-equal? (get-cell-value "D1" xlsx) 100)
      )))))

(run-tests test-test5)
