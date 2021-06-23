#lang racket

(require rackunit/text-ui)

(require "../../../writer.rkt")
(require "../../../lib/lib.rkt")
(require "../../../sheet/sheet.rkt")

(require rackunit "../../../xlsx/xlsx.rkt")

(define test-add-style
  (test-suite
   "test-add-style"

   (test-case
    "test-add-style"

    (parameterize 
     ([*CURRENT_XLSX* (new-xlsx)])

     (add-data-sheet "Sheet1" '(("chenxiao" "love" "陈思衡")))
     (with-sheet
      "Sheet1"
      (lambda ()           
        (add-cell-style "A1" '((fontSize . 20) (fontName . "Impact")))))
     
     (let ([style_index->hash_map (XLSX-style_index->hash_map (*CURRENT_XLSX*))]
           [font_style_index->hash_map (XLSX-font_style_index->hash_map (*CURRENT_XLSX*))])
       (check-equal? (hash-count style_index->hash_map) 1)
       (check-equal? (hash-ref style_index->hash_map 1)
                     (make-hash '((fontSize . 20) (fontName . "Impact"))))

       (check-equal? (hash-count font_style_index->hash_map) 1)
       (check-equal? (hash-ref font_style_index->hash_map 1)
                     (make-hash '((fontSize . 20) (fontName . "Impact")))))
     
     (with-sheet
      "Sheet1"
      (lambda ()
        (let ([cell->style_index_map (DATA-SHEET-cell->style_index_map (*CURRENT_SHEET*))])
          (check-equal? (hash-count cell->style_index_map) 1)
          (check-equal? (hash-ref cell->style_index_map "A1") 1))))

     ))))
    
(run-tests test-add-style)
