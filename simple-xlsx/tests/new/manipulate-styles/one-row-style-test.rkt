#lang racket

(require rackunit/text-ui)

(require "../../../writer.rkt")
(require "../../../lib/lib.rkt")
(require "../../../sheet/sheet.rkt")

(require rackunit "../../../xlsx/xlsx.rkt")

(define test-one-row-style
  (test-suite
   "test-one-row-style"

   (test-case
    "test-one-row-style"

    (parameterize 
     ([*CURRENT_XLSX* (new-xlsx)])

     (add-data-sheet "Sheet1" '(("chenxiao" "love" "陈思衡")))

     (with-sheet
      "Sheet1"
      (lambda ()           
        (add-cell-style "A1-C3" '((fontSize . 20) (fontName . "Impact")))
        (add-cell-style "B2-D4" '((fontSize . 21)))
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
              [cell->style_index_map (DATA-SHEET-cell->style_index_map (*CURRENT_SHEET*))]
              [sheet_row->cells_map (DATA-SHEET-row->cells_map (*CURRENT_SHEET*))]
              )
          (check-equal? (hash-count cell->style_index_map) 14)
          (check-equal? (set-count (hash-ref sheet_row->cells_map 
          (check-equal? (hash-ref cell->style_index_map "A1") 1)
          (check-equal? (hash-ref cell->style_index_map "A3") 1)
          (check-equal? (hash-ref cell->style_index_map "C1") 1)
          (check-equal? (hash-ref cell->style_index_map "B2") 2)
          (check-equal? (hash-ref cell->style_index_map "C2") 2)
          (check-equal? (hash-ref cell->style_index_map "B3") 2)
          (check-equal? (hash-ref cell->style_index_map "C3") 2)
          (check-equal? (hash-ref cell->style_index_map "B4") 3)
          (check-equal? (hash-ref cell->style_index_map "D2") 3)
          (check-equal? (hash-ref cell->style_index_map "D4") 3)
          )))
     ))
   ))
    
(run-tests test-one-row-style)
