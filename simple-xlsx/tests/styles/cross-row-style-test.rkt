#lang racket

(require rackunit/text-ui)

(require rackunit "../../styles.rkt")

(require "../../writer.rkt")
(require "../../lib/lib.rkt")
(require "../../sheet/sheet.rkt")
(require "../../xlsx/xlsx.rkt")

(define test-cross-row-style
  (test-suite
   "test-cross-row-style"

   (test-case
    "test-cross-row-style"

    (parameterize 
     ([*CURRENT_XLSX* (new-xlsx)])

     (add-data-sheet "Sheet1" '(("chenxiao" "love" "陈思衡")))

     (with-sheet
      "Sheet1"
      (lambda ()           
        (add-row-style "1-3" '((fontSize . 20) (fontName . "Impact")))
        (add-row-style "2-4" '((fontSize . 21)))
        ))
     
     (let (
           [style_index->hash_map (XLSX-style_index->hash_map (*CURRENT_XLSX*))]
           [font_style_index->hash_map (XLSX-font_style_index->hash_map (*CURRENT_XLSX*))]
           [num_style_index->hash_map (XLSX-num_style_index->hash_map (*CURRENT_XLSX*))]
           )

       (check-equal? (hash-count style_index->hash_map) 3)
       (check-equal? (hash-ref style_index->hash_map 1) (make-hash '((fontSize . 20) (fontName . "Impact"))))
       (check-equal? (hash-ref style_index->hash_map 2) (make-hash '((fontSize . 21) (fontName . "Impact"))))
       (check-equal? (hash-ref style_index->hash_map 3) (make-hash '((fontSize . 21))))
       
       (check-equal? (hash-count font_style_index->hash_map) 3)
       (check-equal? (hash-ref font_style_index->hash_map 1) (make-hash '((fontSize . 20) (fontName . "Impact"))))
       (check-equal? (hash-ref font_style_index->hash_map 2) (make-hash '((fontSize . 21) (fontName . "Impact"))))
       (check-equal? (hash-ref font_style_index->hash_map 3) (make-hash '((fontSize . 21))))
       )
     
     (with-sheet
      "Sheet1"
      (lambda ()
        (let (
              [row->style_index_map (DATA-SHEET-row->style_index_map (*CURRENT_SHEET*))]
              [sheet_row->cells_map (DATA-SHEET-row->cells_map (*CURRENT_SHEET*))]
              )
          (check-equal? (hash-count row->style_index_map) 4)
          (check-equal? (hash-ref row->style_index_map 1) 1)
          (check-equal? (hash-ref row->style_index_map 2) 2)
          (check-equal? (hash-ref row->style_index_map 3) 2)
          (check-equal? (hash-ref row->style_index_map 4) 3)
          (check-equal? (set-count (hash-ref sheet_row->cells_map 1 '())) 0)
          )))
     ))
   ))
    
(run-tests test-cross-row-style)
