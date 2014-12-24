#lang racket

(provide test-test5)

(require rackunit/text-ui)

(require rackunit "../../main.rkt")

(require "../../lib/lib.rkt")

(define test-test5
  (test-suite
   "test-test5"

   (with-input-from-xlsx-file
    (build-path "test5" "test5.xlsx")
    (lambda ()
      (test-case
       "test-get-sheet-data"

       (load-sheet "Sheet1")
       (check-equal? (get-cell-value "D1") 100)
      )))))
