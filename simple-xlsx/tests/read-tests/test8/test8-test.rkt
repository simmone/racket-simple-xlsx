#lang racket

(require rackunit/text-ui)

(require rackunit)

(require "../../../main.rkt")

(require racket/runtime-path)
(define-runtime-path test_file "test8.xlsx")

(define test-test8
  (test-suite
   "test-test8"

   (with-input-from-xlsx-file
    test_file
    (lambda (xlsx)
      (test-case
       "test-get-sheet-data"

       (load-sheet (car (get-sheet-names xlsx)) xlsx)
       (check-equal? (get-cell-value "D3" xlsx) "20160930")
       (check-equal? (get-cell-value "H15" xlsx) "EB1221075")
      )))))

(run-tests test-test8)
