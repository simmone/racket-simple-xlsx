#lang racket

(require rackunit/text-ui)

(require rackunit)

(require "../../../main.rkt")

(require racket/runtime-path)
(define-runtime-path test_file "test13.xlsx")

(define test-test13
  (test-suite
   "test-test13"

   (test-case
    "test-get-cell-value"
    (with-input-from-xlsx-file
     test_file
     (lambda (xlsx)
       (load-sheet-ref 0 xlsx)
       (check-equal? (get-cell-value "A1" xlsx) 2)
       )))

   (test-case
    "test-get-cell-formula"
    (with-input-from-xlsx-file
     test_file
     (lambda (xlsx)
       (load-sheet-ref 0 xlsx)
       (check-equal? (get-cell-formula "A1" xlsx) "1+1")
       )))
    ))

(run-tests test-test13)
