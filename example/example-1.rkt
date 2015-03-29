#lang racket

(require rackunit/text-ui)

(require rackunit "../simple-xlsx/main.rkt")

(let ([xlsx (new xlsx-data%)])
  (let ([col_attr_hash (make-hash)])
    (hash-set! col_attr_hash "A" (col-attr 100 "red"))
    (hash-set! col_attr_hash "B" (col-attr 200 "blue"))
    (send xlsx add-sheet '(("chenxiao" "陈晓") (1 2 34 100 456.34)) "Sheet1" #:col_attr_hash col_attr_hash))
  (send xlsx add-sheet '((1 2 3 4)) "Sheet2")
  (send xlsx add-sheet '(()) "Sheet3")
  (write-xlsx-file xlsx "example1.xlsx"))
  
