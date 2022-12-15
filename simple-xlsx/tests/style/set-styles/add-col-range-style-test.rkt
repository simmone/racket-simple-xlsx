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
    "test-add-col-style-range"

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
          (set-col-range-font-style "1" 10 "Arial" "0000ff")
          (set-col-range-alignment-style "1" "center" "center")
          (set-col-range-number-style "E" "0.000")

          (check-equal? (hash-count (SHEET-STYLE-col->style_map (*CURRENT_SHEET_STYLE*))) 2)
          (check-equal? (STYLE-hash_code (hash-ref (SHEET-STYLE-col->style_map (*CURRENT_SHEET_STYLE*)) 1))
                        "<s>10<p>Arial<p>0000FF<s>center<p>center<s><s>")
          (check-equal? (STYLE-hash_code (hash-ref (SHEET-STYLE-col->style_map (*CURRENT_SHEET_STYLE*)) 5))
                        "<s><s><s>0.000<s>")

          (check-equal? (hash-count (SHEET-STYLE-cell->style_map (*CURRENT_SHEET_STYLE*))) 4)
          (check-equal? (STYLE-hash_code (hash-ref (SHEET-STYLE-cell->style_map (*CURRENT_SHEET_STYLE*)) "A1"))
                        "<s>10<p>Arial<p>0000FF<s>center<p>center<s><s>")
          (check-equal? (STYLE-hash_code (hash-ref (SHEET-STYLE-cell->style_map (*CURRENT_SHEET_STYLE*)) "A2"))
                        "<s>10<p>Arial<p>0000FF<s>center<p>center<s><s>")
          (check-false (hash-has-key? (SHEET-STYLE-cell->style_map (*CURRENT_SHEET_STYLE*)) "B1"))
          (check-false (hash-has-key? (SHEET-STYLE-cell->style_map (*CURRENT_SHEET_STYLE*)) "B2"))
          (check-false (hash-has-key? (SHEET-STYLE-cell->style_map (*CURRENT_SHEET_STYLE*)) "C2"))
          (check-false (hash-has-key? (SHEET-STYLE-cell->style_map (*CURRENT_SHEET_STYLE*)) "C2"))
          (check-false (hash-has-key? (SHEET-STYLE-cell->style_map (*CURRENT_SHEET_STYLE*)) "D2"))
          (check-false (hash-has-key? (SHEET-STYLE-cell->style_map (*CURRENT_SHEET_STYLE*)) "D2"))
          (check-equal? (STYLE-hash_code (hash-ref (SHEET-STYLE-cell->style_map (*CURRENT_SHEET_STYLE*)) "E1"))
                        "<s><s><s>0.000<s>")
          (check-equal? (STYLE-hash_code (hash-ref (SHEET-STYLE-cell->style_map (*CURRENT_SHEET_STYLE*)) "E2"))
                        "<s><s><s>0.000<s>")
          )))))
    ))

(run-tests test-styles)
