#lang racket

(require rackunit/text-ui)

(require rackunit "../../styles.rkt")

(require "../../writer.rkt")
(require "../../lib/lib.rkt")
(require "../../sheet/sheet.rkt")
(require "../../xlsx/xlsx.rkt")

(define test-one-col-style
  (test-suite
   "test-one-col-style"

   (test-case
    "test-one-col-style"

    (parameterize 
     ([*CURRENT_XLSX* (new-xlsx)])

     (add-data-sheet "Sheet1" '(("chenxiao" "love" "陈思衡")))

     (with-sheet
      "Sheet1"
      (lambda ()           
        (add-col-style "A-C" '((fontSize . 20) (fontName . "Impact")))
        ))
     
     (let (
           [style_index->hash_map (XLSX-style_index->hash_map (*CURRENT_XLSX*))]
           [font_style_index->hash_map (XLSX-font_style_index->hash_map (*CURRENT_XLSX*))]
           [num_style_index->hash_map (XLSX-num_style_index->hash_map (*CURRENT_XLSX*))]
           )

       (check-equal? (hash-count style_index->hash_map) 1)
       (check-equal? (hash-ref style_index->hash_map 1) (make-hash '((fontSize . 20) (fontName . "Impact"))))
       
       (check-equal? (hash-count font_style_index->hash_map) 1)
       (check-equal? (hash-ref font_style_index->hash_map 1) (make-hash '((fontSize . 20) (fontName . "Impact"))))
       )
     
     (with-sheet
      "Sheet1"
      (lambda ()
        (let (
              [col->style_index_map (DATA-SHEET-col->style_index_map (*CURRENT_SHEET*))]
              [sheet_col->cells_map (DATA-SHEET-col->cells_map (*CURRENT_SHEET*))]
              )
          (check-equal? (hash-count col->style_index_map) 3)
          (check-equal? (hash-ref col->style_index_map 1) 1)
          (check-equal? (hash-ref col->style_index_map 2) 1)
          (check-equal? (hash-ref col->style_index_map 3) 1)
          (check-equal? (set-count (hash-ref sheet_col->cells_map 1 '())) 0)
          )))
     ))
   ))
    
(run-tests test-one-col-style)
