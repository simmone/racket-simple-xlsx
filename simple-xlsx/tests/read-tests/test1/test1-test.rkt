#lang racket

(require rackunit)
(require rackunit/text-ui)

(require racket/runtime-path)
(define-runtime-path test_file "test1.xlsx")

(require "../../../main.rkt")

(define test-test1
  (test-suite
   "test-test1"

   (with-input-from-xlsx-file
    test_file
    (lambda (xlsx)
      (test-case 
       "test-get-sheets"
       
       (check-equal? (get-sheet-names xlsx) '("Sheet1" "Sheet2" "Sheet3")))
   
      (test-case
       "test-get-sheet-data"
       
       (load-sheet "Sheet1" xlsx)
       (check-equal? (get-cell-value "A1" xlsx) "chenxiao")
       (check-equal? (get-cell-value "B1" xlsx) "love")
       (check-equal? (get-cell-value "C1" xlsx) "chensiheng")
       (check-equal? (get-cell-value "C2" xlsx) 7)
       )

      (test-case
       "test-load-sheet-ref"
       
       (load-sheet-ref 0 xlsx)
       (check-equal? (get-cell-value "A1" xlsx) "chenxiao")
       (check-equal? (get-cell-value "B1" xlsx) "love")
       (check-equal? (get-cell-value "C1" xlsx) "chensiheng")
       (check-equal? (get-cell-value "C2" xlsx) 7)
       )

      (test-case
       "test-get-sheet-dimension"

       (let ([dimension (get-sheet-dimension xlsx)])
         (check-equal? (car dimension) 2)
         (check-equal? (cdr dimension) 3))
       )

      (test-case
       "test-get-sheet-rows"
       (let ([rows (get-sheet-rows xlsx)])
         (check-equal? (first (list-ref rows 0)) "chenxiao")
         (check-equal? (second (list-ref rows 0)) "love")
         (check-equal? (third (list-ref rows 0)) "chensiheng")
         (check-equal? (first (list-ref rows 1)) 2)
         (check-equal? (second (list-ref rows 1)) 5)
         (check-equal? (third (list-ref rows 1)) 7)
         ))

      (test-case
       "test-sheet-name-rows"
       (let ([rows (sheet-name-rows test_file "Sheet1")])
         (check-equal? (first (list-ref rows 0)) "chenxiao")
         (check-equal? (second (list-ref rows 0)) "love")
         (check-equal? (third (list-ref rows 0)) "chensiheng")
         (check-equal? (first (list-ref rows 1)) 2)
         (check-equal? (second (list-ref rows 1)) 5)
         (check-equal? (third (list-ref rows 1)) 7)
         ))

      (test-case
       "test-sheet-ref-rows"
       (let ([rows (sheet-ref-rows test_file 0)])
         (check-equal? (first (list-ref rows 0)) "chenxiao")
         (check-equal? (second (list-ref rows 0)) "love")
         (check-equal? (third (list-ref rows 0)) "chensiheng")
         (check-equal? (first (list-ref rows 1)) 2)
         (check-equal? (second (list-ref rows 1)) 5)
         (check-equal? (third (list-ref rows 1)) 7)
         ))

      ))))

(run-tests test-test1)
