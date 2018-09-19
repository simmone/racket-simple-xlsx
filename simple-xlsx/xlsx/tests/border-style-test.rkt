#lang racket

(require rackunit/text-ui)

(require rackunit "../xlsx.rkt")
(require rackunit "../sheet.rkt")

(define test-xlsx
  (test-suite
   "test-xlsx"
   
   (test-case
    "test-border-style"

    (let ([xlsx (new xlsx%)])

      (send xlsx add-data-sheet #:sheet_name "测试1" #:sheet_data 
       '((1 2 "chenxiao") (3 4 "xiaomin") (5 6 "chenxiao") (1 "xx" "simmone")))

      (let* ([sheet (sheet-content (send xlsx get-sheet-by-name "测试1"))])

        (send xlsx add-data-sheet-cell-style! #:sheet_name "测试1" #:cell_range "A1-A4" #:style '( (borderDirection . "all") (borderColor . "red") (borderStyle . "dashed")))
        (let* ([cell_to_origin_style_hash (data-sheet-cell_to_origin_style_hash sheet)])
          (check-equal? (hash-count cell_to_origin_style_hash) 4))

        (send xlsx add-data-sheet-cell-style! #:sheet_name "测试1" #:cell_range "B1-B4" #:style '( (borderDirection . "left") (borderColor . "red") (borderStyle . "dashed")))
        (let* ([cell_to_origin_style_hash (data-sheet-cell_to_origin_style_hash sheet)])
          (check-equal? (hash-count cell_to_origin_style_hash) 8))

        (send xlsx add-data-sheet-cell-style! #:sheet_name "测试1" #:cell_range "C1-C4" #:style '( (borderDirection . "right") (borderColor . "red") (borderStyle . "thick")))
        (let* ([cell_to_origin_style_hash (data-sheet-cell_to_origin_style_hash sheet)])
          (check-equal? (hash-count cell_to_origin_style_hash) 12))
        
        (send xlsx burn-styles!)

        (let* ([xlsx_style (get-field style xlsx)]
               [cell_to_style_index_hash (data-sheet-cell_to_style_index_hash sheet)]
               [style_list (xlsx-style-style_list xlsx_style)]
               [border_code_to_border_index_hash (xlsx-style-border_code_to_border_index_hash xlsx_style)]
               [border_list (xlsx-style-border_list xlsx_style)]
              )

          (check-equal? (hash-count cell_to_style_index_hash) 12)
          (check-equal? (length style_list) 3)
          (check-equal? (hash-count border_code_to_border_index_hash) 3)
          (check-equal? (length border_list) 3)
          
          (check-equal? 
           (list-ref border_list (sub1 (hash-ref (list-ref style_list (sub1 (hash-ref cell_to_style_index_hash "A1"))) 'border)))
           (make-hash '((borderDirection . "all") (borderColor . "red") (borderStyle . "dashed"))))

          (check-equal? 
           (list-ref border_list (sub1 (hash-ref (list-ref style_list (sub1 (hash-ref cell_to_style_index_hash "C3"))) 'border)))
           (make-hash '((borderDirection . "right") (borderColor . "red") (borderStyle . "thick"))))

          (check-equal? 
           (list-ref border_list (sub1 (hash-ref (list-ref style_list (sub1 (hash-ref cell_to_style_index_hash "B1"))) 'border)))
           (make-hash '((borderDirection . "left") (borderColor . "red") (borderStyle . "dashed"))))
          )

        )))))

(run-tests test-xlsx)
