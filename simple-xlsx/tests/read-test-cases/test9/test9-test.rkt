#lang racket

(require rackunit/text-ui)

(require rackunit)

(require "../../../main.rkt")

(require racket/runtime-path)
(define-runtime-path test_file "test9.xlsx")

(define test-test9
  (test-suite
   "test-test9"

   (with-input-from-xlsx-file
    test_file
    (lambda (xlsx)
      (test-case
       "test-get-sheet-data"

       (load-sheet-ref 0 xlsx)
       (check-equal? (get-cell-value "A1" xlsx) 1)
      )))))

(run-tests test-test9)
