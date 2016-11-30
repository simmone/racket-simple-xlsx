#lang racket

(require rackunit/text-ui)

(require rackunit "xlsx.rkt")

(define test-xlsx
  (test-suite
   "test-xlsx"
   
   (test-case
    "test-check-data-list"
    
    (check-exn exn:fail? (lambda () (check-equal? (check-data-list '()) #f)))
    (check-exn exn:fail? (lambda () ((check-data-list '((1) 4)))))
    (check-exn exn:fail? (lambda () (check-data-list '((1) (1 2)))))
    
    (check-true (check-data-list '((1 2) (3 4))))
    )

   (test-case
    "test-add-data-sheet-string-item-map"

    (let ([xlsx (new xlsx%)])
      (send xlsx add-data-sheet "测试1" '((1 2 "chenxiao") (3 4 "xiaomin") (5 6 "chenxiao") (1 "xx" "simmone")))
      
      (let ([string_item_map (get-field string_item_map xlsx)])
        (check-equal? (hash-count string_item_map) 4)
        (check-true (hash-has-key? string_item_map "xx")))))

   (test-case
    "test-xlsx"

    (let ([xlsx (new xlsx%)])
      (check-equal? (get-field sheets xlsx) '())
      
      (send xlsx add-data-sheet "测试1" '((1)))

      (check-exn exn:fail? (lambda () (send xlsx add-data-sheet "测试1" '())))

      (send xlsx add-data-sheet "测试2" '((1)))
      
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

        (send xlsx set-data-sheet-col-width! "测试2" "A-C" 100)
        (check-equal? (hash-ref (data-sheet-width_hash (sheet-content sheet)) "A-C") 100)

        (send xlsx set-data-sheet-cell-color! "测试2" "A1-C2" "red")
        (check-equal? (hash-ref (data-sheet-color_hash (sheet-content sheet)) "A1-C2") "red")
        )

      (send xlsx add-line-chart-sheet "测试3" "图表1")

      (send xlsx add-line-chart-sheet "测试4" "图表2")

      (check-exn exn:fail? (lambda () (send xlsx add-data-sheet "测试1" '())))
      (check-exn exn:fail? (lambda () (send xlsx add-line-chart-sheet "测试4" "test")))

      (send xlsx add-data-sheet "测试5" '((1 2 3 4) (4 5 6 7) (8 9 10 11)))

      (let ([sheet (send xlsx get-sheet-by-name "测试3")])
        (check-equal? (sheet-name sheet) "测试3"))

      (check-exn exn:fail? (lambda () (check-data-range-valid xlsx "测试5" "E1-E3")))

      (check-exn exn:fail? (lambda () (check-data-range-valid xlsx "测试5" "C1-C4")))

      (check-equal? (send xlsx get-range-data "测试5" "A1-A3") '(1 4 8))
      (check-equal? (send xlsx get-range-data "测试5" "B1-B3") '(2 5 9))
      (check-equal? (send xlsx get-range-data "测试5" "C1-C3") '(3 6 10))
      (check-equal? (send xlsx get-range-data "测试5" "D1-D3") '(4 7 11))

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

        (send xlsx set-line-chart-x-data! "测试4" "测试5" "A1-A3")
        (let ([data_range (line-chart-sheet-x_data_range (sheet-content sheet))])
          (check-equal? (data-range-range_str data_range) "A1-A3")
          (check-equal? (data-range-sheet_name data_range) "测试5"))
        
        (send xlsx add-line-chart-y-data! "测试4" "折线1" "测试5" "B1-B3")

        (send xlsx add-line-chart-y-data! "测试4" "折线2" "测试5" "C1-C3")
        
        (let* ([y_data_list (line-chart-sheet-y_data_range_list (sheet-content sheet))]
               [y_data1 (first y_data_list)]
               [y_data2 (second y_data_list)])
          (check-equal? (data-serial-topic y_data1) "折线1")
          (check-equal? (data-range-sheet_name (data-serial-data_range y_data1)) "测试5")
          (check-equal? (data-range-range_str (data-serial-data_range y_data1)) "B1-B3")

          (check-equal? (data-serial-topic y_data2) "折线2")
          (check-equal? (data-range-sheet_name (data-serial-data_range y_data2)) "测试5")
          (check-equal? (data-range-range_str (data-serial-data_range y_data2)) "C1-C3")
        ))
      ))

   (test-case
    "test-convert-range"
    
    (check-equal? (convert-range "C2-C10") "$C$2:$C$10")

    (check-equal? (convert-range "AB20-AB100") "$AB$20:$AB$100")
    )
   
   (test-case
    "test-check-range"
    
    (check-exn exn:fail? (lambda () (check-range "c2")))
    (check-exn exn:fail? (lambda () (check-range "c2-c2")))

    (check-exn exn:fail? (lambda () (check-range "A2-A1")))
    (check-exn exn:fail? (lambda () (check-range "A2-B3")))
    )
   
   (test-case
    "test-check-col-range"
    
    (check-col-range "A-Z")
    
    (check-exn exn:fail? (lambda () (check-col-range "B-A")))

    (check-exn exn:fail? (lambda () (check-col-range "A1-A")))
    )

   (test-case
    "test-check-cell-range"
    
    (check-cell-range "A1-B2")

    (check-exn exn:fail? (lambda () (check-cell-range "A10-B9")))

    (check-exn exn:fail? (lambda () (check-cell-range "B1-A1")))
    )

   (test-case
    "test-range-length"
    
    (check-equal? (range-length "$A$2:$A$20") 19)
    (check-equal? (range-length "$AB$21:$AB$21") 1)
    )

   (test-case
    "test-set-data-sheet-cell-color-and-get-style-list"

    (let ([xlsx (new xlsx%)])
      (send xlsx add-data-sheet "测试1" '((1 2 "chenxiao") (3 4 "xiaomin") (5 6 "chenxiao") (1 "xx" "simmone")))

      (send xlsx set-data-sheet-cell-color! "测试1" "A1-A4" "red")

      (send xlsx set-data-sheet-cell-color! "测试1" "B1-B4" "blue")
      
      (check-equal? (send xlsx get-styles-list) '("blue" "red"))
      )

      )

   ))

(run-tests test-xlsx)
