#lang racket

(require rackunit/text-ui)

(require rackunit "xlsx.rkt")

(define test-xlsx
  (test-suite
   "test-xlsx"

   (test-case
    "test-xlsx"

    (let ([xlsx (new xlsx%)])
      (check-equal? (get-field sheets xlsx) '())
      
      (send xlsx add-data-sheet "测试" '())

      (send xlsx add-data-sheet "测试1" '(1))
      
      (let ([sheet (send xlsx sheet-ref 0)])
        (check-equal? (sheet-name sheet) "测试")
        (check-equal? (sheet-seq sheet) 1)
        (check-equal? (sheet-type sheet) 'data)
        (check-equal? (sheet-typeSeq sheet) 1)
        )
      
      (let ([sheet (send xlsx sheet-ref 1)])
        (check-equal? (sheet-name sheet) "测试1")
        (check-equal? (sheet-seq sheet) 2)
        (check-equal? (sheet-type sheet) 'data)
        (check-equal? (sheet-typeSeq sheet) 2)

        (send xlsx set-sheet-col-width! sheet "A-C" 100)
        (check-equal? (hash-ref (data-sheet-width_hash (sheet-content sheet)) "A-C") 100)

        (send xlsx set-sheet-col-color! sheet "A-C" "red")
        (check-equal? (hash-ref (data-sheet-color_hash (sheet-content sheet)) "A-C") "red")
        )
      )
    )
   ))

(run-tests test-xlsx)
