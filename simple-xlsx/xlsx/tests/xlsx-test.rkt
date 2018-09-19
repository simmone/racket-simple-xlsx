#lang racket

(require rackunit/text-ui)
(require racket/date)

(require rackunit "../xlsx.rkt")
(require rackunit "../sheet.rkt")

(define test-xlsx
  (test-suite
   "test-xlsx"
   
   (test-case
    "test-add-data-sheet-string-item-map"

    (let ([xlsx (new xlsx%)])
      (send xlsx add-data-sheet 
            #:sheet_name "测试1" 
            #:sheet_data '((1 2 "chenxiao") (3 4 "xiaomin") (5 6 "chenxiao") (1 "xx" "simmone")))
      
      (let ([string_item_map (get-field string_item_map xlsx)])
        (check-equal? (hash-count string_item_map) 4)
        (check-true (hash-has-key? string_item_map "xx")))))

   (test-case
    "test-add-data-sheet-date-type"

    (let ([xlsx (new xlsx%)])
      (send xlsx add-data-sheet 
            #:sheet_name "测试1" 
            #:sheet_data (list
                          (list 1 (seconds->date (find-seconds 0 0 0 17 9 2018)))
                          (list 2 (seconds->date (find-seconds 0 0 0 18 9 2018)))
                          (list 3 (seconds->date (find-seconds 0 0 0 19 9 2018)))
                          ))
      
      (let* ([sheets (get-field sheets xlsx)]
             [data_sheet (data-sheet-rows (sheet-content (first sheets)))])
        (check-equal? (second (list-ref data_sheet 0)) 43360)
        (check-equal? (second (list-ref data_sheet 1)) 43361)
        (check-equal? (second (list-ref data_sheet 2)) 43362))))

   (test-case
    "test-xlsx"

    (let ([xlsx (new xlsx%)])
      (check-equal? (get-field sheets xlsx) '())
      
      (send xlsx add-data-sheet #:sheet_name "测试1" #:sheet_data '((1)))

      (check-exn exn:fail? (lambda () (send xlsx add-data-sheet #:sheet_name "测试1" #:sheet_data '())))

      (send xlsx add-data-sheet #:sheet_name "测试2" #:sheet_data '((1)))
      
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

        (send xlsx set-data-sheet-col-width! #:sheet_name "测试2" #:col_range "A-C" #:width 100)
        (check-equal? (hash-ref (data-sheet-width_hash (sheet-content sheet)) "A-C") 100)
        )

      (send xlsx add-chart-sheet #:sheet_name "测试3" #:topic "图表1")

      (send xlsx add-chart-sheet #:sheet_name "测试4" #:topic "图表2" #:x_topic "万元")

      (check-exn exn:fail? (lambda () (send xlsx add-data-sheet #:sheet_name "测试1" #:sheet_data '())))
      (check-exn exn:fail? (lambda () (send xlsx add-chart-sheet #:sheet_name "测试4" #:topic "test")))

      (send xlsx add-data-sheet #:sheet_name "测试5" #:sheet_data '((1 2 3 4) (4 5 6 7) (8 9 10 11)))

      (let ([sheet (send xlsx get-sheet-by-name "测试3")])
        (check-equal? (sheet-name sheet) "测试3"))

      (check-exn exn:fail? (lambda () (send xlsx check-data-range-valid #:sheet_name "测试5" #:range_str "E1-E3")))

      (check-exn exn:fail? (lambda () (send xlsx check-data-range-valid #:sheet_name "测试5" #:range_str "C1-C4")))

      (check-equal? (send xlsx get-range-data "测试5" "A1-A3") '(1 4 8))
      (check-equal? (send xlsx get-range-data "测试5" "B1-B3") '(2 5 9))
      (check-equal? (send xlsx get-range-data "测试5" "C1-C3") '(3 6 10))
      (check-equal? (send xlsx get-range-data "测试5" "D1-D3") '(4 7 11))
      (check-equal? (send xlsx get-range-data "测试5" "A1-D1") '(1 2 3 4))
      (check-equal? (send xlsx get-range-data "测试5" "A2-D2") '(4 5 6 7))
      (check-equal? (send xlsx get-range-data "测试5" "A3-D3") '(8 9 10 11))

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

        (check-equal? (chart-sheet-topic (sheet-content sheet)) "图表2")
        (check-equal? (chart-sheet-x_topic (sheet-content sheet)) "万元")

        (send xlsx set-chart-x-data! #:sheet_name "测试4" #:data_sheet_name "测试5" #:data_range "A1-A3")
        (let ([data_range (chart-sheet-x_data_range (sheet-content sheet))])

          (check-equal? (data-range-range_str data_range) "A1-A3")
          (check-equal? (data-range-sheet_name data_range) "测试5")
          )
        
        (send xlsx add-chart-serial! #:sheet_name "测试4" #:data_sheet_name "测试5" #:y_topic "折线1" #:data_range "B1-B3")

        (send xlsx add-chart-serial! #:sheet_name "测试4" #:data_sheet_name "测试5" #:y_topic "折线2" #:data_range "C1-C3")
        
        (let* ([y_data_list (chart-sheet-y_data_range_list (sheet-content sheet))]
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
    "test-get-string-index-map"

    (let ([xlsx (new xlsx%)])
      (send xlsx add-data-sheet #:sheet_name "测试1" #:sheet_data '(("C" "D" "B") ("B" "Z" "A")))
      
      (let ([string_index_map (send xlsx get-string-index-map)])
        (check-equal? (hash-count string_index_map) 5)
        (check-equal? (hash-ref string_index_map "A") 0)
        (check-equal? (hash-ref string_index_map "B") 1)
        (check-equal? (hash-ref string_index_map "C") 2)
        (check-equal? (hash-ref string_index_map "D") 3)
        (check-equal? (hash-ref string_index_map "Z") 4)
      )
      ))

   ))

(run-tests test-xlsx)
