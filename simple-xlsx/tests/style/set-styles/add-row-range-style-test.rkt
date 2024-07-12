#lang racket

(require rackunit/text-ui
         rackunit
         "../../../xlsx/xlsx.rkt"
         "../../../sheet/sheet.rkt"
         "../../../style/style.rkt"
         "../../../style/border-style.rkt"
         "../../../style/font-style.rkt"
         "../../../style/alignment-style.rkt"
         "../../../style/number-style.rkt"
         "../../../style/fill-style.rkt"
         "../../../style/styles.rkt"
         "../../../style/set-styles.rkt")

(define test-styles
  (test-suite
   "test-styles"

   (test-case
    "test-add-cell-style-range"

    (with-xlsx
     (lambda ()
       (add-data-sheet "Sheet1"
                       '(("month1" "month2" "month3" "month4" "real") (201601 100 110 1110 6.9)))
       (add-data-sheet "Sheet2" '((1)))
       (add-data-sheet "Sheet3" '((1)))
       (add-chart-sheet "Chart1" 'LINE "Chart1" '())
       (add-chart-sheet "Chart2" 'LINE "Chart2" '())
       (add-chart-sheet "Chart3" 'LINE "Chart3" '())

       (with-sheet
        (lambda ()
          (set-row-range-font-style "1" 10 "Arial" "0000FF")
          (set-row-range-alignment-style "1" "center" "center")
          (set-row-range-number-style "2" "0.000")
          (check-equal? (hash-count (SHEET-STYLE-row->style_map (*CURRENT_SHEET_STYLE*))) 2)
          (check-equal? (hash-ref (SHEET-STYLE-row->style_map (*CURRENT_SHEET_STYLE*)) 1)
                        (STYLE
                          #f (FONT-STYLE 10 "Arial" "0000FF") (ALIGNMENT-STYLE "center" "center") #f #f))
          (check-equal? (hash-ref (SHEET-STYLE-row->style_map (*CURRENT_SHEET_STYLE*)) 2)
                        (STYLE #f #f #f (NUMBER-STYLE "1" "0.000") #f))
          (check-equal? (hash-count (SHEET-STYLE-cell->style_map (*CURRENT_SHEET_STYLE*))) 10)
          (check-equal? (hash-ref (SHEET-STYLE-cell->style_map (*CURRENT_SHEET_STYLE*)) "A1")
                        (STYLE
                          #f (FONT-STYLE 10 "Arial" "0000FF") (ALIGNMENT-STYLE "center" "center") #f #f))
          (check-equal? (hash-ref (SHEET-STYLE-cell->style_map (*CURRENT_SHEET_STYLE*)) "B1")
                        (STYLE
                          #f (FONT-STYLE 10 "Arial" "0000FF") (ALIGNMENT-STYLE "center" "center") #f #f))
          (check-equal? (hash-ref (SHEET-STYLE-cell->style_map (*CURRENT_SHEET_STYLE*)) "C1")
                        (STYLE
                          #f (FONT-STYLE 10 "Arial" "0000FF") (ALIGNMENT-STYLE "center" "center") #f #f))
          (check-equal? (hash-ref (SHEET-STYLE-cell->style_map (*CURRENT_SHEET_STYLE*)) "D1")
                        (STYLE
                          #f (FONT-STYLE 10 "Arial" "0000FF") (ALIGNMENT-STYLE "center" "center") #f #f))
          (check-equal? (hash-ref (SHEET-STYLE-cell->style_map (*CURRENT_SHEET_STYLE*)) "E1")
                        (STYLE
                          #f (FONT-STYLE 10 "Arial" "0000FF") (ALIGNMENT-STYLE "center" "center") #f #f))
          (check-equal? (hash-ref (SHEET-STYLE-cell->style_map (*CURRENT_SHEET_STYLE*)) "A2")
                        (STYLE
                          #f #f #f (NUMBER-STYLE "1" "0.000") #f))
          (check-equal? (hash-ref (SHEET-STYLE-cell->style_map (*CURRENT_SHEET_STYLE*)) "B2")
                        (STYLE
                          #f #f #f (NUMBER-STYLE "1" "0.000") #f))
          (check-equal? (hash-ref (SHEET-STYLE-cell->style_map (*CURRENT_SHEET_STYLE*)) "C2")
                        (STYLE
                          #f #f #f (NUMBER-STYLE "1" "0.000") #f))
          (check-equal? (hash-ref (SHEET-STYLE-cell->style_map (*CURRENT_SHEET_STYLE*)) "D2")
                        (STYLE
                          #f #f #f (NUMBER-STYLE "1" "0.000") #f))
          (check-equal? (hash-ref (SHEET-STYLE-cell->style_map (*CURRENT_SHEET_STYLE*)) "E2")
                        (STYLE
                          #f #f #f (NUMBER-STYLE "1" "0.000") #f))
          )))))
    ))

(run-tests test-styles)
