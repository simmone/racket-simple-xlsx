#lang racket

(require rackunit/text-ui)

(require rackunit)

(require "../../../main.rkt")

(require racket/runtime-path)
(define-runtime-path test_file "test12.xlsx")

(define test-test12
  (test-suite
   "test-test12"

   (test-case
    "test-get-cell-value"
    (with-input-from-xlsx-file
     test_file
     (lambda (xlsx)
       (load-sheet-ref 0 xlsx)
       (check-equal? (get-cell-value "A2" xlsx) "20000000040931")
       )))
    ))

(run-tests test-test12)
