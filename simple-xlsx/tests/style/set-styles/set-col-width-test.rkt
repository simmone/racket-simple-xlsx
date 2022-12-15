#lang racket

(require rackunit/text-ui rackunit)

(require"../../../xlsx/xlsx.rkt")
(require"../../../sheet/sheet.rkt")
(require"../../../style/style.rkt")
(require"../../../style/set-styles.rkt")

(define test-styles
  (test-suite
   "test-styles"

   (test-case
    "test-styles-set-col-width"

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
          (set-col-range-width "1-3" 5)
          (set-col-range-width "C-F" 6)
          (set-col-range-width "G-10" 7)
          (set-col-range-width "G-10" 8)

          (check-equal? (hash-count (SHEET-STYLE-col->width_map (*CURRENT_SHEET_STYLE*))) 10)
          (check-equal? (hash-count (SHEET-STYLE-col->width_map (hash-ref (STYLES-sheet_index->style_map (*STYLES*)) 0))) 10)
          (check-equal? (hash-count (SHEET-STYLE-col->width_map (*CURRENT_SHEET_STYLE*))) 10)

          (check-equal? (hash-ref (SHEET-STYLE-col->width_map (*CURRENT_SHEET_STYLE*)) 1) 5)
          (check-equal? (hash-ref (SHEET-STYLE-col->width_map (*CURRENT_SHEET_STYLE*)) 2) 5)

          (check-equal? (hash-ref (SHEET-STYLE-col->width_map (*CURRENT_SHEET_STYLE*)) 3) 6)
          (check-equal? (hash-ref (SHEET-STYLE-col->width_map (*CURRENT_SHEET_STYLE*)) 4) 6)
          (check-equal? (hash-ref (SHEET-STYLE-col->width_map (*CURRENT_SHEET_STYLE*)) 5) 6)
          (check-equal? (hash-ref (SHEET-STYLE-col->width_map (*CURRENT_SHEET_STYLE*)) 6) 6)

          (check-equal? (hash-ref (SHEET-STYLE-col->width_map (*CURRENT_SHEET_STYLE*)) 7) 8)
          (check-equal? (hash-ref (SHEET-STYLE-col->width_map (*CURRENT_SHEET_STYLE*)) 8) 8)
          (check-equal? (hash-ref (SHEET-STYLE-col->width_map (*CURRENT_SHEET_STYLE*)) 9) 8)
          (check-equal? (hash-ref (SHEET-STYLE-col->width_map (*CURRENT_SHEET_STYLE*)) 10) 8)
          )))))
   ))

(run-tests test-styles)
