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
          (set-cell-range-border-style "A1" "top" "0000ff" "dashed")
          (set-cell-range-border-style "A1" "bottom" "f000ff" "thin")
          (set-cell-range-border-style "A1" "left" "ff00ff" "double")
          (set-cell-range-border-style "A1" "right" "fff0ff" "thick")
          (set-cell-range-font-style "B1" 10 "Arial" "0000ff")
          (set-cell-range-alignment-style "C1" "center" "center")
          (set-cell-range-number-style "D1" "0.000")
          (set-cell-range-fill-style "E1" "FF0000" "solid")
          (check-equal? (hash-count (SHEET-STYLE-cell->style_map (*CURRENT_SHEET_STYLE*))) 5)
          (check-equal? (STYLE-hash_code (hash-ref (SHEET-STYLE-cell->style_map (*CURRENT_SHEET_STYLE*)) "A1"))
                        "0000FF<p>dashed<p>F000FF<p>thin<p>FF00FF<p>double<p>FFF0FF<p>thick<s><s><s><s>")
          (check-equal? (STYLE-hash_code (hash-ref (SHEET-STYLE-cell->style_map (*CURRENT_SHEET_STYLE*)) "B1"))
                        "<s>10<p>Arial<p>0000FF<s><s><s>")
          (check-equal? (STYLE-hash_code (hash-ref (SHEET-STYLE-cell->style_map (*CURRENT_SHEET_STYLE*)) "C1"))
                        "<s><s>center<p>center<s><s>")
          (check-equal? (STYLE-hash_code (hash-ref (SHEET-STYLE-cell->style_map (*CURRENT_SHEET_STYLE*)) "D1"))
                        "<s><s><s>0.000<s>")
          (check-equal? (STYLE-hash_code (hash-ref (SHEET-STYLE-cell->style_map (*CURRENT_SHEET_STYLE*)) "E1"))
                        "<s><s><s><s>FF0000<p>solid")

          (set-cell-range-border-style "A1-C2" "all" "0000ff" "dashed")
          (set-cell-range-font-style "A1-C2" 10 "Arial" "0000ff")
          (set-cell-range-alignment-style "A1-C2" "center" "center")
          (set-cell-range-number-style "A1-C2" "0.000")
          (set-cell-range-fill-style "A1-C2" "FF0000" "solid")
          (check-equal? (hash-count (SHEET-STYLE-cell->style_map (*CURRENT_SHEET_STYLE*))) 8)
          (check-equal?
           (STYLE-hash_code (hash-ref (SHEET-STYLE-cell->style_map (*CURRENT_SHEET_STYLE*)) "A1"))
           "0000FF<p>dashed<p>0000FF<p>dashed<p>0000FF<p>dashed<p>0000FF<p>dashed<s>10<p>Arial<p>0000FF<s>center<p>center<s>0.000<s>FF0000<p>solid")
          (check-equal?
           (STYLE-hash_code (hash-ref (SHEET-STYLE-cell->style_map (*CURRENT_SHEET_STYLE*)) "C2"))
           "0000FF<p>dashed<p>0000FF<p>dashed<p>0000FF<p>dashed<p>0000FF<p>dashed<s>10<p>Arial<p>0000FF<s>center<p>center<s>0.000<s>FF0000<p>solid")
          ))))
    )))

(run-tests test-styles)
