#lang racket

(require rackunit/text-ui)

(require rackunit "../../main.rkt")

(require "../../lib/lib.rkt")

(define test-test1
  (test-suite
   "test-test1"

   (let ([data_list '((("chenxiao" "陈晓") (1 2 34 100 456.34)) ((1 2 3 4)) ())])
     (dynamic-wind
         (lambda ()
           (write-xlsx-file data_list #f "test1.xlsx"))
         (lambda ()
           (with-input-from-xlsx-file
            "test1.xlsx"
            (lambda ()
              (test-case 
               "test-get-sheets"
       
               (check-equal? (get-sheet-names) '("Sheet1" "Sheet2" "Sheet3")))
   
              (test-case
               "test-get-sheet-data"
       
               (load-sheet "Sheet1")
               (check-equal? (get-cell-value "A1") "chenxiao")
               (check-equal? (get-cell-value "B1") "陈晓")
               (check-equal? (get-cell-value "E2") 456.34)))))
         (lambda ()
           (delete-file "test1.xlsx"))))))

(run-tests test-test1)
