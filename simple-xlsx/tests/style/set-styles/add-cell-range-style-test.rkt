#lang racket

(require rackunit/text-ui
         rackunit
         "../../../xlsx/xlsx.rkt"
         "../../../sheet/sheet.rkt"
         "../../../style/border-style.rkt"
         "../../../style/font-style.rkt"
         "../../../style/alignment-style.rkt"
         "../../../style/number-style.rkt"
         "../../../style/fill-style.rkt"
         "../../../style/style.rkt"
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
       (add-data-sheet "Sheet2"
                       '(("month1" "month2" "month3" "month4" "real") (201601 100 110 1110 6.9)))
       (add-data-sheet "Sheet3" '((1)))
       (add-chart-sheet "Chart1" 'LINE "Chart1" '())
       (add-chart-sheet "Chart2" 'LINE "Chart2" '())
       (add-chart-sheet "Chart3" 'LINE "Chart3" '())

       (with-sheet
        (lambda ()
          (set-cell-range-border-style "A1" "top" "0000FF" "dashed")
          (set-cell-range-border-style "A1" "bottom" "F000FF" "thin")
          (set-cell-range-border-style "A1" "left" "FF00FF" "double")
          (set-cell-range-border-style "A1" "right" "FFF0FF" "thick")
          (set-cell-range-font-style "B1" 10 "Arial" "0000FF")
          (set-cell-range-alignment-style "C1" "center" "center")
          (set-cell-range-number-style "D1" "0.000")
          (set-cell-range-fill-style "E1" "FF0000" "solid")
          (check-equal? (hash-count (SHEET-STYLE-cell->style_map (*CURRENT_SHEET_STYLE*))) 5)
          (check-equal? (hash-ref (SHEET-STYLE-cell->style_map (*CURRENT_SHEET_STYLE*)) "A1")
                        (STYLE
                         (BORDER-STYLE "FF00FF" "double" "FFF0FF" "thick" "0000FF" "dashed" "F000FF" "thin")
                         #f #f #f #f))
          (check-equal? (hash-ref (SHEET-STYLE-cell->style_map (*CURRENT_SHEET_STYLE*)) "B1")
                        (STYLE
                         #f (FONT-STYLE 10 "Arial" "0000FF") #f #f #f))
          (check-equal? (hash-ref (SHEET-STYLE-cell->style_map (*CURRENT_SHEET_STYLE*)) "C1")
                        (STYLE
                         #f #f (ALIGNMENT-STYLE "center" "center") #f #f))
          (check-equal? (hash-ref (SHEET-STYLE-cell->style_map (*CURRENT_SHEET_STYLE*)) "D1")
                        (STYLE
                         #f #f #f (NUMBER-STYLE "1" "0.000") #f))
          (check-equal? (hash-ref (SHEET-STYLE-cell->style_map (*CURRENT_SHEET_STYLE*)) "E1")
                        (STYLE
                         #f #f #f #f (FILL-STYLE "FF0000" "solid")))

          (set-cell-range-border-style "A1-C2" "all" "0000FF" "dashed")
          (set-cell-range-font-style "A1-C2" 10 "Arial" "0000FF")
          (set-cell-range-alignment-style "A1-C2" "center" "center")
          (set-cell-range-number-style "A1-C2" "0.000")
          (set-cell-range-fill-style "A1-C2" "FF0000" "solid")
          (check-equal? (hash-count (SHEET-STYLE-cell->style_map (*CURRENT_SHEET_STYLE*))) 8)
          (check-equal?
           (hash-ref (SHEET-STYLE-cell->style_map (*CURRENT_SHEET_STYLE*)) "A1")
           (STYLE
            (BORDER-STYLE "0000FF" "dashed" "0000FF" "dashed" "0000FF" "dashed" "0000FF" "dashed")
            (FONT-STYLE 10 "Arial" "0000FF")
            (ALIGNMENT-STYLE "center" "center")
            (NUMBER-STYLE "1" "0.000")
            (FILL-STYLE "FF0000" "solid")))
          (check-equal?
           (hash-ref (SHEET-STYLE-cell->style_map (*CURRENT_SHEET_STYLE*)) "C2")
           (STYLE
            (BORDER-STYLE "0000FF" "dashed" "0000FF" "dashed" "0000FF" "dashed" "0000FF" "dashed")
            (FONT-STYLE 10 "Arial" "0000FF")
            (ALIGNMENT-STYLE "center" "center")
            (NUMBER-STYLE "1" "0.000")
            (FILL-STYLE "FF0000" "solid")))
          ))

       (with-sheet-ref
        1
        (lambda ()
          (set-cell-range-border-style "A1-B2" "all" "0000FF" "dashed")
          (check-equal?
           (hash-ref (SHEET-STYLE-cell->style_map (*CURRENT_SHEET_STYLE*)) "A1")
           (STYLE
            (BORDER-STYLE "0000FF" "dashed" "0000FF" "dashed" "0000FF" "dashed" "0000FF" "dashed") #f #f #f #f))
          (check-equal?
           (hash-ref (SHEET-STYLE-cell->style_map (*CURRENT_SHEET_STYLE*)) "A2")
           (STYLE
            (BORDER-STYLE "0000FF" "dashed" "0000FF" "dashed" "0000FF" "dashed" "0000FF" "dashed") #f #f #f #f))
          (set-cell-range-border-style "B1-C2" "all" "0000FF" "thin")
          (check-equal?
           (hash-ref (SHEET-STYLE-cell->style_map (*CURRENT_SHEET_STYLE*)) "A1")
           (STYLE
            (BORDER-STYLE "0000FF" "dashed" "0000FF" "dashed" "0000FF" "dashed" "0000FF" "dashed") #f #f #f #f))
          (check-equal?
           (hash-ref (SHEET-STYLE-cell->style_map (*CURRENT_SHEET_STYLE*)) "A2")
           (STYLE
            (BORDER-STYLE "0000FF" "dashed" "0000FF" "dashed" "0000FF" "dashed" "0000FF" "dashed") #f #f #f #f))
          (check-equal?
           (hash-ref (SHEET-STYLE-cell->style_map (*CURRENT_SHEET_STYLE*)) "B1")
           (STYLE
            (BORDER-STYLE "0000FF" "thin" "0000FF" "thin" "0000FF" "thin" "0000FF" "thin") #f #f #f #f))
          (check-equal?
           (hash-ref (SHEET-STYLE-cell->style_map (*CURRENT_SHEET_STYLE*)) "B2")
           (STYLE
            (BORDER-STYLE "0000FF" "thin" "0000FF" "thin" "0000FF" "thin" "0000FF" "thin") #f #f #f #f))
          (check-equal?
           (hash-ref (SHEET-STYLE-cell->style_map (*CURRENT_SHEET_STYLE*)) "C1")
           (STYLE
            (BORDER-STYLE "0000FF" "thin" "0000FF" "thin" "0000FF" "thin" "0000FF" "thin") #f #f #f #f))
          (check-equal?
           (hash-ref (SHEET-STYLE-cell->style_map (*CURRENT_SHEET_STYLE*)) "C2")
           (STYLE
            (BORDER-STYLE "0000FF" "thin" "0000FF" "thin" "0000FF" "thin" "0000FF" "thin") #f #f #f #f))
          ))

       )))
   ))

(run-tests test-styles)
