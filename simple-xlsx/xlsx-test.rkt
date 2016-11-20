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
      
      (send xlsx add-data-sheet "测试1" '())

      (send xlsx add-data-sheet "测试2" '(1))
      
      (let ([sheet (send xlsx sheet-ref 0)])
        (check-equal? (sheet-name sheet) "测试1")
        (check-equal? (sheet-seq sheet) 1)
        (check-equal? (sheet-type sheet) 'data)
        (check-equal? (sheet-typeSeq sheet) 1)
        )
      
      (let ([sheet (send xlsx sheet-ref 1)])
        (check-equal? (sheet-name sheet) "测试2")
        (check-equal? (sheet-seq sheet) 2)
        (check-equal? (sheet-type sheet) 'data)
        (check-equal? (sheet-typeSeq sheet) 2)

        (send xlsx set-sheet-col-width! sheet "A-C" 100)
        (check-equal? (hash-ref (data-sheet-width_hash (sheet-content sheet)) "A-C") 100)

        (send xlsx set-sheet-col-color! sheet "A-C" "red")
        (check-equal? (hash-ref (data-sheet-color_hash (sheet-content sheet)) "A-C") "red")
        )

      (send xlsx add-line-chart-sheet "测试3" "图表1")

      (send xlsx add-line-chart-sheet "测试4" "图表2")

      (let ([sheet (send xlsx sheet-ref 2)])
        (check-equal? (sheet-name sheet) "测试3")
        (check-equal? (sheet-seq sheet) 3)
        (check-equal? (sheet-type sheet) 'chart)
        (check-equal? (sheet-typeSeq sheet) 1)
        )
      
      (let ([sheet (send xlsx sheet-ref 3)])
        (check-equal? (sheet-name sheet) "测试4")
        (check-equal? (sheet-seq sheet) 4)
        (check-equal? (sheet-type sheet) 'chart)
        (check-equal? (sheet-typeSeq sheet) 2)

        (check-equal? (line-chart-sheet-topic (sheet-content sheet)) "图表2")

        (send xlsx set-line-chart-x-data! (sheet-content sheet) '(1 2 3 4))
        (check-equal? (line-chart-sheet-x_data (sheet-content sheet)) '(1 2 3 4))
        
        (send xlsx add-line-chart-y-data! (sheet-content sheet) "折线1" '(5 6 7 8))
        (send xlsx add-line-chart-y-data! (sheet-content sheet) "折线2" '(15 16 17 18))
        (check-equal? (data-serial-topic (first (line-chart-sheet-y_data_list (sheet-content sheet)))) "折线1")
        (check-equal? (data-serial-data_list (first (line-chart-sheet-y_data_list (sheet-content sheet)))) '(5 6 7 8))
        (check-equal? (data-serial-topic (second (line-chart-sheet-y_data_list (sheet-content sheet)))) "折线2")
        (check-equal? (data-serial-data_list (second (line-chart-sheet-y_data_list (sheet-content sheet)))) '(15 16 17 18))
        )
      )
    )
   ))

(run-tests test-xlsx)
