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

      (let ([sheet (send xlsx get-sheet-by-name "测试3")])
        (check-equal? (sheet-name sheet) "测试3"))

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

        (send xlsx set-line-chart-x-data! (sheet-content sheet) "测试1" "A2-A10")
        (let ([data_range (line-chart-sheet-x_data_range (sheet-content sheet))])
          (check-equal? (data-range-range_str data_range) "A2-A10")
          (check-equal? (data-range-sheet_name data_range) "测试1"))
        
        (send xlsx add-line-chart-y-data! (sheet-content sheet) "折线1" "测试1" "B2-B10")

        (send xlsx add-line-chart-y-data! (sheet-content sheet) "折线2" "测试2" "C2-C10")

        (let* ([y_data_list (line-chart-sheet-y_data_range_list (sheet-content sheet))]
               [y_data1 (first y_data_list)]
               [y_data2 (second y_data_list)])
          (check-equal? (data-serial-topic y_data1) "折线1")
          (check-equal? (data-range-sheet_name (data-serial-data_range y_data1)) "测试1")
          (check-equal? (data-range-range_str (data-serial-data_range y_data1)) "B2-B10")

          (check-equal? (data-serial-topic y_data2) "折线2")
          (check-equal? (data-range-sheet_name (data-serial-data_range y_data2)) "测试2")
          (check-equal? (data-range-range_str (data-serial-data_range y_data2)) "C2-C10")
        ))
      ))

   (test-case
    "test-convert-range"
    
    (check-equal? (convert-range "C2-C10") "$C$2:$C$10")

    (check-equal? (convert-range "AB20-AB100") "$AB$20:$AB$100")
    )

   (test-case
    "test-range-length"
    
    (check-equal? (range-length "$A$2:$A$20") 19)
    (check-equal? (range-length "$AB$21:$AB$21") 1)
    )

   ))

(run-tests test-xlsx)
