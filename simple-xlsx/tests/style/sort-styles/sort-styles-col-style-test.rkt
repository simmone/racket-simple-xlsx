#lang racket

(require rackunit/text-ui rackunit)

(require"../../../xlsx/xlsx.rkt")
(require"../../../sheet/sheet.rkt")
(require"../../../style/style.rkt")
(require"../../../style/sort-styles.rkt")
(require"../../../style/set-styles.rkt")

(define test-styles
  (test-suite
   "test-styles"

   (test-case
    "test-sort-styles"

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
          (check-equal? (hash-count (STYLES-style->index_map (*STYLES*))) 0)
          (check-equal? (hash-count (SHEET-STYLE-col->style_map (*CURRENT_SHEET_STYLE*))) 0)

          (set-col-range-font-style "1" 10 "Arial" "0000ff")
          ))

       (sort-styles)

       (check-equal? (hash-count (STYLES-style->index_map (*STYLES*))) 1)
       (check-equal? (hash-ref (STYLES-style->index_map (*STYLES*)) "<s>10<p>Arial<p>0000FF<s><s><s>") 0)

       )))
   ))

(run-tests test-styles)
