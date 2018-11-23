#lang racket

(require rackunit/text-ui)

(require rackunit)

(require "../../../main.rkt")

(require racket/runtime-path)
(define-runtime-path test_file "test11.xlsx")

(define test-test11
  (test-suite
   "test-test11"

   (test-case
    "test-get-cell-value"
    (with-input-from-xlsx-file
     test_file
     (lambda (xlsx)
       (load-sheet-ref 0 xlsx)
       (check-equal? (get-cell-value "AG2" xlsx) 1000000000)
       )))
    ))

(run-tests test-test11)
