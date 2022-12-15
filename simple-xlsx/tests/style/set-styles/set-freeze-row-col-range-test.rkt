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
    "test-styles-set-freeze-row-col-range"

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
          (set-freeze-row-col-range 1 3)

          (check-equal? (SHEET-STYLE-freeze_range (*CURRENT_SHEET_STYLE*)) '(1 . 3))
          (check-equal? (SHEET-STYLE-freeze_range (hash-ref (STYLES-sheet_index->style_map (*STYLES*)) 0)) '(1 . 3))
          ))

       (with-sheet-ref
        1
        (lambda ()
          (set-freeze-row-col-range 2 4)

          (check-equal? (SHEET-STYLE-freeze_range (*CURRENT_SHEET_STYLE*)) '(2 . 4))
          (check-equal? (SHEET-STYLE-freeze_range (hash-ref (STYLES-sheet_index->style_map (*STYLES*)) 1)) '(2 . 4))
          ))

       (check-equal? (SHEET-STYLE-freeze_range (hash-ref (STYLES-sheet_index->style_map (*STYLES*)) 0)) '(1 . 3))
       (check-equal? (SHEET-STYLE-freeze_range (hash-ref (STYLES-sheet_index->style_map (*STYLES*)) 1)) '(2 . 4))
       )))
   ))

(run-tests test-styles)
