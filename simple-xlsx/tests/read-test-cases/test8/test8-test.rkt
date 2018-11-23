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

       (load-sheet-ref 0 xlsx)
       (check-equal? (get-cell-value "A1" xlsx) "日本語の漢字がダメかもしれません。")
       (check-equal? (get-cell-value "A2" xlsx) "漢字がダメかもしれません。")
       (check-equal? (get-cell-value "A3" xlsx) "かんじがダメかもしれません。")
      )))))

(run-tests test-test8)
