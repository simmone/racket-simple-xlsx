#lang racket

(require rackunit/text-ui)

(require rackunit "../xlsx.rkt")
(require rackunit "../sheet.rkt")

(define test-xlsx
  (test-suite
   "test-xlsx"
   
   (test-case
    "test-fill-style"

    (let ([xlsx (new xlsx%)])

      (send xlsx add-data-sheet #:sheet_name "测试1" #:sheet_data 
       '((1 2 "chenxiao") (3 4 "xiaomin") (5 6 "chenxiao") (1 "xx" "simmone")))

      (let* ([sheet (sheet-content (send xlsx get-sheet-by-name "测试1"))])

        (send xlsx add-data-sheet-cell-style! #:sheet_name "测试1" #:cell_range "A1-A4" #:style '( (backgroundColor . "FF0000") ))
        (let* ([cell_to_origin_style_hash (data-sheet-cell_to_origin_style_hash sheet)])
          (check-equal? (hash-count cell_to_origin_style_hash) 4))

        (send xlsx add-data-sheet-cell-style! #:sheet_name "测试1" #:cell_range "B1-B4" #:style '( (backgroundColor . "blue") ))
        (let* ([cell_to_origin_style_hash (data-sheet-cell_to_origin_style_hash sheet)])
          (check-equal? (hash-count cell_to_origin_style_hash) 8))

        (send xlsx add-data-sheet-cell-style! #:sheet_name "测试1" #:cell_range "C1-C4" #:style '( (backgroundColor . "FF0000") ))
        (let* ([cell_to_origin_style_hash (data-sheet-cell_to_origin_style_hash sheet)])
          (check-equal? (hash-count cell_to_origin_style_hash) 12))
        
        (send xlsx burn-styles!)

        (let* ([xlsx_style (get-field style xlsx)]
               [cell_to_style_index_hash (data-sheet-cell_to_style_index_hash sheet)]
               [style_list (xlsx-style-style_list xlsx_style)]
               [fill_code_to_fill_index_hash (xlsx-style-fill_code_to_fill_index_hash xlsx_style)]
               [fill_list (xlsx-style-fill_list xlsx_style)]
              )

          (check-equal? (hash-count cell_to_style_index_hash) 12)
          (check-equal? (length style_list) 2)
          (check-equal? (hash-count fill_code_to_fill_index_hash) 2)
          (check-equal? (length fill_list) 2)
          
          (check-equal? 
           (list-ref fill_list (- (hash-ref (list-ref style_list (sub1 (hash-ref cell_to_style_index_hash "A1"))) 'fill) 2))
           (make-hash '((fgColor . "FF0000"))))

          (check-equal? 
           (list-ref fill_list (- (hash-ref (list-ref style_list (sub1 (hash-ref cell_to_style_index_hash "B1"))) 'fill) 2))
           (make-hash '((fgColor . "blue"))))

          (check-equal? 
           (list-ref fill_list (- (hash-ref (list-ref style_list (sub1 (hash-ref cell_to_style_index_hash "C1"))) 'fill) 2))
           (make-hash '((fgColor . "FF0000"))))
          )

        (let ([style_map (send xlsx get-cell-to-style-index-map "测试1")])
          (check-equal? (hash-count style_map) 12)
          (check-equal? (hash-ref style_map "A1") 1)
          (check-equal? (hash-ref style_map "B2") 2)
          (check-equal? (hash-ref style_map "C2") 1)
          )

        )))))

(run-tests test-xlsx)
