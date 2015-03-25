#lang racket

(require rackunit/text-ui)

(require rackunit "../../main.rkt")

(require "../../lib/lib.rkt")

(define test-test1
  (test-suite
   "test-test1"

   (dynamic-wind
       (lambda ()
         (let ([xlsx (new xlsx-data%)])
           (send xlsx add-sheet '(("chenxiao" "陈晓") (1 2 34 100 456.34)) "Sheet1")
           (send xlsx add-sheet '((1 2 3 4)) "Sheet2")
           (send xlsx add-sheet '(()) "Sheet3")
           (write-xlsx-file xlsx "test1.xlsx")))
       (lambda ()
         (with-input-from-xlsx-file "test1.xlsx"
                                    (lambda (xlsx)
                                      (test-case 
                                       "test-get-sheets"
                                       
                                       (check-equal? (get-sheet-names xlsx) '("Sheet1" "Sheet2" "Sheet3")))
                                      
                                      (test-case
                                       "test-get-sheet-data"
                                       
                                       (load-sheet "Sheet1" xlsx)
                                       (check-equal? (get-cell-value "A1" xlsx) "chenxiao")
                                       (check-equal? (get-cell-value "B1" xlsx) "陈晓")
                                       (check-equal? (get-cell-value "E2" xlsx) 456.34)))))
       (lambda ()
         (delete-file "test1.xlsx")))))

(run-tests test-test1)
