#lang racket

(require rackunit/text-ui)

(require rackunit "../../styles.rkt")

(require "../../writer.rkt")
(require "../../lib/lib.rkt")
(require "../../sheet/sheet.rkt")
(require "../../xlsx/xlsx.rkt")

(define test-one-cell-style
  (test-suite
   "test-one-cell-style"

   (test-case
    "test-one-cell-style"

    (parameterize 
     ([*CURRENT_XLSX* (new-xlsx)])

     (add-data-sheet "Sheet1" '(("chenxiao" "love" "陈思衡")))

     (with-sheet
      "Sheet1"
      (lambda ()           
        (add-cell-style "A1" '((fontSize . 20) (fontName . "Impact")))
        (add-cell-style "B1" '((fontSize . 20) (fontName . "Impact")))
        (add-cell-style "B1" '((fontSize . 21)))
        (add-cell-style "C1" '((fontSize . 21)))
        (add-cell-style "D1" '((numberPrecision . 2)))
        (add-cell-style "E1" '((fontSize . 20) (fontName . "Impact")))
        (add-cell-style "E1" '((numberPrecision . 2)))
        (add-cell-style "F1" '((fontName . "Impact") (numberPrecision . 2)))
        (add-cell-style "G1" '((fontSize . 20) (fontName . "Impact") (numberPrecision . 2)))
        ))
     
     (let (
           [style_index->hash_map (XLSX-style_index->hash_map (*CURRENT_XLSX*))]
           [font_style_index->hash_map (XLSX-font_style_index->hash_map (*CURRENT_XLSX*))]
           [num_style_index->hash_map (XLSX-num_style_index->hash_map (*CURRENT_XLSX*))]
           )

       (check-equal? (hash-count style_index->hash_map) 6)
       (check-equal? (hash-ref style_index->hash_map 1) (make-hash '((fontSize . 20) (fontName . "Impact"))))
       (check-equal? (hash-ref style_index->hash_map 2) (make-hash '((fontSize . 21) (fontName . "Impact"))))
       (check-equal? (hash-ref style_index->hash_map 3) (make-hash '((fontSize . 21))))
       (check-equal? (hash-ref style_index->hash_map 4) (make-hash '((numberPrecision . 2))))
       (check-equal? (hash-ref style_index->hash_map 5) (make-hash '((fontSize . 20) (fontName . "Impact") (numberPrecision . 2))))
       (check-equal? (hash-ref style_index->hash_map 6) (make-hash '((fontName . "Impact") (numberPrecision . 2))))
       
       (check-equal? (hash-count font_style_index->hash_map) 4)
       (check-equal? (hash-ref font_style_index->hash_map 1) (make-hash '((fontSize . 20) (fontName . "Impact"))))
       (check-equal? (hash-ref font_style_index->hash_map 2) (make-hash '((fontSize . 21) (fontName . "Impact"))))
       (check-equal? (hash-ref font_style_index->hash_map 3) (make-hash '((fontSize . 21))))
       (check-equal? (hash-ref font_style_index->hash_map 4) (make-hash '((fontName . "Impact"))))

       (check-equal? (hash-count num_style_index->hash_map) 1)
       (check-equal? (hash-ref num_style_index->hash_map 1) (make-hash '((numberPrecision . 2))))
       )
     
     (with-sheet
      "Sheet1"
      (lambda ()
        (let ([cell->style_index_map (DATA-SHEET-cell->style_index_map (*CURRENT_SHEET*))])
          (check-equal? (hash-count cell->style_index_map) 7)
          (check-equal? (hash-ref cell->style_index_map "A1") 1)
          (check-equal? (hash-ref cell->style_index_map "B1") 2)
          (check-equal? (hash-ref cell->style_index_map "C1") 3)
          (check-equal? (hash-ref cell->style_index_map "D1") 4)
          (check-equal? (hash-ref cell->style_index_map "E1") 5)
          (check-equal? (hash-ref cell->style_index_map "F1") 6)
          (check-equal? (hash-ref cell->style_index_map "G1") 5)
          )))
     ))
   ))
    
(run-tests test-one-cell-style)
