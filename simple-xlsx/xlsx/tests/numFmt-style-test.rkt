#lang racket

(require rackunit/text-ui)

(require rackunit "../xlsx.rkt")
(require rackunit "../sheet.rkt")

(define test-xlsx
  (test-suite
   "test-xlsx"
   
   (test-case
    "test-numFmt-style"

    (let ([xlsx (new xlsx%)])

      (send xlsx add-data-sheet #:sheet_name "测试1" #:sheet_data 
       '((1 2 "chenxiao" 43360) (3 4 "xiaomin" 43361) (5 6 "chenxiao" 43362) (1 "xx" "simmone" 43363)))

      (let* ([sheet (sheet-content (send xlsx get-sheet-by-name "测试1"))])

        (send xlsx add-data-sheet-cell-style! #:sheet_name "测试1" #:cell_range "A1-A4" #:style '( (numberPrecision . 2) ))
        (let* ([cell_to_origin_style_hash (data-sheet-cell_to_origin_style_hash sheet)])
          (check-equal? (hash-count cell_to_origin_style_hash) 4))

        (send xlsx add-data-sheet-cell-style! #:sheet_name "测试1" #:cell_range "B1-B4" #:style '( (numberPrecision . 3) ))
        (let* ([cell_to_origin_style_hash (data-sheet-cell_to_origin_style_hash sheet)])
          (check-equal? (hash-count cell_to_origin_style_hash) 8))

        (send xlsx add-data-sheet-cell-style! #:sheet_name "测试1" #:cell_range "C1-C4" #:style '( (numberPrecision . 2) ))
        (let* ([cell_to_origin_style_hash (data-sheet-cell_to_origin_style_hash sheet)])
          (check-equal? (hash-count cell_to_origin_style_hash) 12))

        (send xlsx add-data-sheet-cell-style! #:sheet_name "测试1" #:cell_range "D1-D4" #:style '( (dateFormat . "yyyy年mm月dd日") ))
        (let* ([cell_to_origin_style_hash (data-sheet-cell_to_origin_style_hash sheet)])
          (check-equal? (hash-count cell_to_origin_style_hash) 16))
        
        (send xlsx burn-styles!)

        (let* ([xlsx_style (get-field style xlsx)]
               [cell_to_style_index_hash (data-sheet-cell_to_style_index_hash sheet)]
               [style_list (xlsx-style-style_list xlsx_style)]
               [numFmt_code_to_numFmt_index_hash (xlsx-style-numFmt_code_to_numFmt_index_hash xlsx_style)]
               [numFmt_list (xlsx-style-numFmt_list xlsx_style)]
              )

          (check-equal? (hash-count cell_to_style_index_hash) 16)
          (check-equal? (length style_list) 3)
          (check-equal? (hash-count numFmt_code_to_numFmt_index_hash) 3)
          (check-equal? (length numFmt_list) 3)
          
          (check-equal? 
           (list-ref numFmt_list 
                     (- (hash-ref (list-ref style_list (sub1 (hash-ref cell_to_style_index_hash "A1"))) 'numFmt) 165))
           (make-hash '((numberPrecision . 2))))

          (check-equal? 
           (list-ref numFmt_list (- (hash-ref (list-ref style_list (sub1 (hash-ref cell_to_style_index_hash "B1"))) 'numFmt) 165))
           (make-hash '((numberPrecision . 3))))

          (check-equal? 
           (list-ref numFmt_list (- (hash-ref (list-ref style_list (sub1 (hash-ref cell_to_style_index_hash "C1"))) 'numFmt) 165))
           (make-hash '((numberPrecision . 2))))

          (check-equal? 
           (list-ref numFmt_list (- (hash-ref (list-ref style_list (sub1 (hash-ref cell_to_style_index_hash "D1"))) 'numFmt) 165))
           (make-hash '((dateFormat . "yyyy年mm月dd日"))))
          )

        (let ([style_map (send xlsx get-cell-to-style-index-map "测试1")])
          (check-equal? (hash-count style_map) 16)
          (check-equal? (hash-ref style_map "A1") 1)
          (check-equal? (hash-ref style_map "B2") 2)
          (check-equal? (hash-ref style_map "C2") 1)
          (check-equal? (hash-ref style_map "D4") 3)
          )

        )))))

(run-tests test-xlsx)
