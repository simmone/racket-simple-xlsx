#lang racket

(require rackunit/text-ui
         rackunit
         "../../../xlsx/xlsx.rkt"
         "../../../sheet/sheet.rkt"
         "../../../style/style.rkt"
         "../../../style/styles.rkt"
         "../../../style/set-styles.rkt")

(define test-styles
  (test-suite
   "test-styles"

   (test-case
    "test-styles-set-row-range-height"

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
          (set-row-range-height "1-3" 5)
          (set-row-range-height "2-4" 6)
          (set-row-range-height "3-5" 7)

          (check-equal? (hash-count (SHEET-STYLE-row->height_map (*CURRENT_SHEET_STYLE*))) 5)
          (check-equal? (hash-count (SHEET-STYLE-row->height_map
                                     (list-ref (STYLES-sheet_style_list (*STYLES*)) 0)))
                        5)
          (check-equal? (hash-count (SHEET-STYLE-row->height_map (*CURRENT_SHEET_STYLE*))) 5)

          (check-equal? (hash-ref (SHEET-STYLE-row->height_map (*CURRENT_SHEET_STYLE*)) 1) 5)
          (check-equal? (hash-ref (SHEET-STYLE-row->height_map (*CURRENT_SHEET_STYLE*)) 2) 6)
          (check-equal? (hash-ref (SHEET-STYLE-row->height_map (*CURRENT_SHEET_STYLE*)) 3) 7)
          (check-equal? (hash-ref (SHEET-STYLE-row->height_map (*CURRENT_SHEET_STYLE*)) 4) 7)
          (check-equal? (hash-ref (SHEET-STYLE-row->height_map (*CURRENT_SHEET_STYLE*)) 5) 7)
          )))))
   ))

(run-tests test-styles)
