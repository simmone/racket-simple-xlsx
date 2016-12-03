#lang racket

(require rackunit/text-ui)

(require rackunit "../../main.rkt")

(define write-test
  (test-suite
   "test-normal-data-sheet"

   (test-case 
    "simple1"
              
    (let ([xlsx (new xlsx%)])
      (send xlsx add-data-sheet "Sheet1" '((1 2 3 4 5)))

      (write-xlsx-file xlsx "test1.xlsx")
      ))
   ))

(run-tests write-test)
