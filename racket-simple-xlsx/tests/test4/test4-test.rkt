#lang racket

(provide test-test4)

(require rackunit/text-ui)

(require rackunit "../../main.rkt")

(require "../../lib/lib.rkt")

(define test-test4
  (test-suite
   "test-test4"

   (with-input-from-xlsx-file
    (build-path "test4" "test4.xlsx")
    (lambda ()
      (test-case
       "test-get-sheet-data"

       (load-sheet "Sheet1")
       (check-equal? (get-cell-value "A1") "200008194477601")
       (check-equal? (get-cell-value "B1") "20140425")
       (check-equal? (get-cell-value "C1") "U0298")
       (check-equal? (get-cell-value "D1") 100)
      )))))
