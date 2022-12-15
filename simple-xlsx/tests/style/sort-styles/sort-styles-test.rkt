#lang racket

(require rackunit/text-ui rackunit)

(require"../../../xlsx/xlsx.rkt")
(require"../../../sheet/sheet.rkt")
(require"../../../style/style.rkt")
(require"../../../style/font-style.rkt")
(require"../../../style/fill-style.rkt")
(require"../../../style/alignment-style.rkt")
(require"../../../style/border-style.rkt")
(require"../../../style/number-style.rkt")
(require"../../../style/sort-styles.rkt")
(require"../../../style/set-styles.rkt")

(define test-styles
  (test-suite
   "test-styles"

   (test-case
    "test-style<?"

    (check-true (style<? (style-from-hash-code "<s><s><s><s>FF0000<p>solid") (style-from-hash-code "<s><s><s>FF0000<p>solid<s>")))
    (check-true (style<? (style-from-hash-code "<s><s><s><s>") (style-from-hash-code "<s><s><s><s>FF0000<p>solid")))
    (check-true (style<? (style-from-hash-code "<s><s><s><s>FF0000<p>solid") (style-from-hash-code "<s><s><s>0.000<s>")))
    (check-false (style<?
                  (style-from-hash-code "000000<p>thin<p>0f0000<p>thick<p>00f000<p>double<p>000f00<p>dashed<s><s><s><s>")
                  (style-from-hash-code "<s>0<s><s><s>")))
    (check-false (style<?
                  (style-from-hash-code "000000<p>thin<p>0f0000<p>thick<p>00f000<p>double<p>000f00<p>dashed<s><s><s><s>")
                  (style-from-hash-code "<s><s>0<s><s>")))
    (check-false (style<?
                  (style-from-hash-code "000000<p>thin<p>0f0000<p>thick<p>00f000<p>double<p>000f00<p>dashed<s><s><s><s>")
                  (style-from-hash-code "<s><s><s>0<s>")))
    (check-false (style<? (style-from-hash-code "<s>10<p>Arial<p>0000ff<s><s><s>") (style-from-hash-code "<s><s><s><s>0")))

   (test-case
    "test-style-null?"

    (check-true (style-null? (style-from-hash-code "<s><s><s><s>")))
    (check-false (style-null? (style-from-hash-code "<s><s><s>0.00<s>")))
    )

    (check-true
     (style<?
      (style-from-hash-code "<s><s><s><s>")
      (style-from-hash-code "f00000<p>thin<p>0f0000<p>thick<p>00f000<p>double<p>000f00<p>dashed<s><s><s><s>")))

    (check-true
     (style<?
      (style-from-hash-code "f00000<p>thin<p>0f0000<p>thick<p>00f000<p>double<p>000f00<p>dashed<s><s><s><s>")
      (style-from-hash-code "f00000<p>thin<p>0f0000<p>thick<p>00f000<p>double<p>000f00<p>dashed<s>10<p>Arial<p>0000ff<s><s><s>")))

    (check-true
     (style<?
      (style-from-hash-code "f00000<p>thin<p>0f0000<p>thick<p>00f000<p>double<p>000f00<p>dashed<s>10<p>Arial<p>0000ff<s><s><s>")
      (style-from-hash-code "f00000<p>thin<p>0f0000<p>thick<p>00f000<p>double<p>000f00<p>dashed<s>10<p>Arial<p>0000ff<s>center<p>center<s><s>")))

    (check-true
     (style<?
      (style-from-hash-code "f00000<p>thin<p>0f0000<p>thick<p>00f000<p>double<p>000f00<p>dashed<s>10<p>Arial<p>0000ff<s>center<p>center<s><s>")
      (style-from-hash-code "f00000<p>thin<p>0f0000<p>thick<p>00f000<p>double<p>000f00<p>dashed<s>10<p>Arial<p>0000ff<s>center<p>center<s>0.000<s>")))

    (check-true
     (style<?
      (style-from-hash-code "f00000<p>thin<p>0f0000<p>thick<p>00f000<p>double<p>000f00<p>dashed<s>10<p>Arial<p>0000ff<s>center<p>center<s>0.000<s>")
      (style-from-hash-code "f00000<p>thin<p>0f0000<p>thick<p>00f000<p>double<p>000f00<p>dashed<s>10<p>Arial<p>0000ff<s>center<p>center<s>0.000<s>FF0000<p>solid")))
    )


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
          (check-equal? (hash-count (SHEET-STYLE-cell->style_map (*CURRENT_SHEET_STYLE*))) 0)
          (check-equal? (hash-count (SHEET-STYLE-row->style_map (*CURRENT_SHEET_STYLE*))) 0)
          (check-equal? (hash-count (SHEET-STYLE-col->style_map (*CURRENT_SHEET_STYLE*))) 0)

          (set-cell-range-border-style "A1" "top" "f00000" "thin")
          (set-cell-range-border-style "A1" "bottom" "0f0000" "thick")
          (set-cell-range-border-style "A1" "left" "00f000" "double")
          (set-cell-range-border-style "A1" "right" "000f00" "dashed")
          (set-cell-range-font-style "B1" 10 "Arial" "0000ff")
          (set-cell-range-alignment-style "C1" "center" "center")
          (set-cell-range-font-style "D1" 10 "Arial" "0000ff")
          (set-cell-range-number-style "D1" "0.000")
          (set-cell-range-fill-style "E1" "FF0000" "solid")))

       (sort-styles)

       (check-equal? (hash-count (STYLES-style->index_map (*STYLES*))) 5)
       (check-equal? (hash-count (STYLES-index->style_map (*STYLES*))) 5)
       (check-equal? (hash-ref (STYLES-style->index_map (*STYLES*)) "<s><s><s><s>FF0000<p>solid") 0)
       (check-equal? (hash-ref (STYLES-index->style_map (*STYLES*)) 0) "<s><s><s><s>FF0000<p>solid")
       (check-equal? (hash-ref (STYLES-style->index_map (*STYLES*)) "<s><s>center<p>center<s><s>") 1)
       (check-equal? (hash-ref (STYLES-index->style_map (*STYLES*)) 1) "<s><s>center<p>center<s><s>")
       (check-equal? (hash-ref (STYLES-style->index_map (*STYLES*)) "<s>10<p>Arial<p>0000FF<s><s><s>") 2)
       (check-equal? (hash-ref (STYLES-index->style_map (*STYLES*)) 2) "<s>10<p>Arial<p>0000FF<s><s><s>")
       (check-equal? (hash-ref (STYLES-style->index_map (*STYLES*)) "<s>10<p>Arial<p>0000FF<s><s>0.000<s>") 3)
       (check-equal? (hash-ref (STYLES-index->style_map (*STYLES*)) 3) "<s>10<p>Arial<p>0000FF<s><s>0.000<s>")
       (check-equal? (hash-ref (STYLES-style->index_map (*STYLES*)) "F00000<p>thin<p>0F0000<p>thick<p>00F000<p>double<p>000F00<p>dashed<s><s><s><s>") 4)
       (check-equal? (hash-ref (STYLES-index->style_map (*STYLES*)) 4) "F00000<p>thin<p>0F0000<p>thick<p>00F000<p>double<p>000F00<p>dashed<s><s><s><s>")

       (with-sheet-ref
        0
        (lambda ()
          (check-equal? (hash-count (SHEET-STYLE-cell->style_map (*CURRENT_SHEET_STYLE*))) 5)
          (check-equal? (STYLE-hash_code (hash-ref (SHEET-STYLE-cell->style_map (*CURRENT_SHEET_STYLE*)) "A1"))
                        "F00000<p>thin<p>0F0000<p>thick<p>00F000<p>double<p>000F00<p>dashed<s><s><s><s>")
          (check-equal? (STYLE-hash_code (hash-ref (SHEET-STYLE-cell->style_map (*CURRENT_SHEET_STYLE*)) "B1"))
                        "<s>10<p>Arial<p>0000FF<s><s><s>")
          (check-equal? (STYLE-hash_code (hash-ref (SHEET-STYLE-cell->style_map (*CURRENT_SHEET_STYLE*)) "C1"))
                        "<s><s>center<p>center<s><s>")
          (check-equal? (STYLE-hash_code (hash-ref (SHEET-STYLE-cell->style_map (*CURRENT_SHEET_STYLE*)) "D1"))
                        "<s>10<p>Arial<p>0000FF<s><s>0.000<s>")
          (check-equal? (STYLE-hash_code (hash-ref (SHEET-STYLE-cell->style_map (*CURRENT_SHEET_STYLE*)) "E1"))
                        "<s><s><s><s>FF0000<p>solid")))

       (check-equal? (hash-count (STYLES-border_style->index_map (*STYLES*))) 1)
       (check-equal? (hash-ref (STYLES-border_style->index_map (*STYLES*)) "F00000<p>thin<p>0F0000<p>thick<p>00F000<p>double<p>000F00<p>dashed") 0)

       (check-equal? (hash-count (STYLES-font_style->index_map (*STYLES*))) 1)
       (check-equal? (hash-ref (STYLES-font_style->index_map (*STYLES*)) "10<p>Arial<p>0000FF") 0)

       (check-equal? (hash-count (STYLES-number_style->index_map (*STYLES*))) 1)
       (check-equal? (hash-ref (STYLES-number_style->index_map (*STYLES*)) "0.000") 0)

       (check-equal? (hash-count (STYLES-fill_style->index_map (*STYLES*))) 1)
       (check-equal? (hash-ref (STYLES-fill_style->index_map (*STYLES*)) "FF0000<p>solid") 0)
       )))

   (test-case
    "test-sort-styles-overlap"

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

          (set-cell-range-border-style "A1" "top" "f00000" "thin")
          (set-cell-range-border-style "A1" "bottom" "0f0000" "thick")
          (set-cell-range-border-style "A1" "left" "00f000" "double")
          (set-cell-range-border-style "A1" "right" "000f00" "dashed")

          (set-cell-range-font-style "B1" 10 "Arial" "0000ff")
          (set-cell-range-alignment-style "C1" "center" "center")
          (set-cell-range-number-style "D1" "0.000")
          (set-cell-range-fill-style "E1" "FF0000" "solid")

          (set-cell-range-border-style "A1" "top" "f00000" "thin")
          (set-cell-range-border-style "A1" "bottom" "0f0000" "thick")
          (set-cell-range-border-style "A1" "left" "00f000" "double")
          (set-cell-range-border-style "A1" "right" "00000f" "dashed")

       (sort-styles)

       (check-equal? (hash-count (STYLES-style->index_map (*STYLES*))) 5)
       (check-equal?
        (hash-ref
         (STYLES-style->index_map (*STYLES*))
         "F00000<p>thin<p>0F0000<p>thick<p>00F000<p>double<p>00000F<p>dashed<s><s><s><s>") 4)
       )))))

   (test-case
    "test-sort-styles-overlap-partly"

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

          (set-cell-range-border-style "A1-B2" "all" "0000ff" "dashed")
          (set-cell-range-border-style "B1-C2" "all" "0000ff" "thin")
          (set-cell-range-font-style "B1" 10 "Arial" "0000ff")
          (set-cell-range-font-style "B2" 20 "Arial" "0000ff")
          (set-cell-range-alignment-style "C1" "center" "center")
          (set-cell-range-alignment-style "C2" "center" "center")
          (set-cell-range-number-style "D1" "0.00")
          (set-cell-range-number-style "A2" "0.00%")
          (set-cell-range-fill-style "E1" "FF0000" "solid")
          (set-cell-range-fill-style "D2" "FFF000" "solid")
          (set-cell-range-border-style "A1" "all" "0000f0" "dashed")))

       (sort-styles)

       (check-equal? (hash-count (STYLES-style->index_map (*STYLES*))) 8)
       (check-equal? (hash-ref (STYLES-style->index_map (*STYLES*)) "<s><s><s><s>FF0000<p>solid") 0)
       (check-equal? (hash-ref (STYLES-style->index_map (*STYLES*)) "<s><s><s><s>FFF000<p>solid") 1)
       (check-equal? (hash-ref (STYLES-style->index_map (*STYLES*)) "<s><s><s>0.00<s>") 2)
       (check-equal? (hash-ref (STYLES-style->index_map (*STYLES*))
                               "0000F0<p>dashed<p>0000F0<p>dashed<p>0000F0<p>dashed<p>0000F0<p>dashed<s><s><s><s>") 3)
       (check-equal? (hash-ref (STYLES-style->index_map (*STYLES*))
                               "0000FF<p>dashed<p>0000FF<p>dashed<p>0000FF<p>dashed<p>0000FF<p>dashed<s><s><s>0.00%<s>") 4)
       (check-equal? (hash-ref (STYLES-style->index_map (*STYLES*))
                               "0000FF<p>thin<p>0000FF<p>thin<p>0000FF<p>thin<p>0000FF<p>thin<s><s>center<p>center<s><s>") 5)
       (check-equal? (hash-ref (STYLES-style->index_map (*STYLES*))
                               "0000FF<p>thin<p>0000FF<p>thin<p>0000FF<p>thin<p>0000FF<p>thin<s>10<p>Arial<p>0000FF<s><s><s>") 6)
       (check-equal? (hash-ref (STYLES-style->index_map (*STYLES*))
                               "0000FF<p>thin<p>0000FF<p>thin<p>0000FF<p>thin<p>0000FF<p>thin<s>20<p>Arial<p>0000FF<s><s><s>") 7)

       (with-sheet-ref
        0
        (lambda ()
          (check-equal? (hash-count (SHEET-STYLE-cell->style_map (*CURRENT_SHEET_STYLE*))) 9)
          (check-equal? (STYLE-hash_code (hash-ref (SHEET-STYLE-cell->style_map (*CURRENT_SHEET_STYLE*)) "A1"))
                        "0000F0<p>dashed<p>0000F0<p>dashed<p>0000F0<p>dashed<p>0000F0<p>dashed<s><s><s><s>")
          (check-equal? (STYLE-hash_code (hash-ref (SHEET-STYLE-cell->style_map (*CURRENT_SHEET_STYLE*)) "A2"))
                        "0000FF<p>dashed<p>0000FF<p>dashed<p>0000FF<p>dashed<p>0000FF<p>dashed<s><s><s>0.00%<s>")
          (check-equal? (STYLE-hash_code (hash-ref (SHEET-STYLE-cell->style_map (*CURRENT_SHEET_STYLE*)) "B1"))
                        "0000FF<p>thin<p>0000FF<p>thin<p>0000FF<p>thin<p>0000FF<p>thin<s>10<p>Arial<p>0000FF<s><s><s>")
          (check-equal? (STYLE-hash_code (hash-ref (SHEET-STYLE-cell->style_map (*CURRENT_SHEET_STYLE*)) "B2"))
                        "0000FF<p>thin<p>0000FF<p>thin<p>0000FF<p>thin<p>0000FF<p>thin<s>20<p>Arial<p>0000FF<s><s><s>")
          (check-equal? (STYLE-hash_code (hash-ref (SHEET-STYLE-cell->style_map (*CURRENT_SHEET_STYLE*)) "C1"))
                        "0000FF<p>thin<p>0000FF<p>thin<p>0000FF<p>thin<p>0000FF<p>thin<s><s>center<p>center<s><s>")
          (check-equal? (STYLE-hash_code (hash-ref (SHEET-STYLE-cell->style_map (*CURRENT_SHEET_STYLE*)) "C2"))
                        "0000FF<p>thin<p>0000FF<p>thin<p>0000FF<p>thin<p>0000FF<p>thin<s><s>center<p>center<s><s>")
          (check-equal? (STYLE-hash_code (hash-ref (SHEET-STYLE-cell->style_map (*CURRENT_SHEET_STYLE*)) "D1"))
                        "<s><s><s>0.00<s>")
          (check-equal? (STYLE-hash_code (hash-ref (SHEET-STYLE-cell->style_map (*CURRENT_SHEET_STYLE*)) "D2"))
                        "<s><s><s><s>FFF000<p>solid")
          (check-equal? (STYLE-hash_code (hash-ref (SHEET-STYLE-cell->style_map (*CURRENT_SHEET_STYLE*)) "E1"))
                        "<s><s><s><s>FF0000<p>solid")
          ))

       (check-equal? (hash-count (STYLES-border_style->index_map (*STYLES*))) 3)
       (check-equal?
        (hash-ref
         (STYLES-border_style->index_map (*STYLES*))
         "0000F0<p>dashed<p>0000F0<p>dashed<p>0000F0<p>dashed<p>0000F0<p>dashed") 0)
       (check-equal?
        (hash-ref
         (STYLES-border_style->index_map (*STYLES*))
         "0000FF<p>dashed<p>0000FF<p>dashed<p>0000FF<p>dashed<p>0000FF<p>dashed") 1)
       (check-equal?
        (hash-ref
         (STYLES-border_style->index_map (*STYLES*))
         "0000FF<p>thin<p>0000FF<p>thin<p>0000FF<p>thin<p>0000FF<p>thin") 2)

       (check-equal? (hash-count (STYLES-font_style->index_map (*STYLES*))) 2)
       (check-equal? (hash-ref (STYLES-font_style->index_map (*STYLES*)) "10<p>Arial<p>0000FF") 0)
       (check-equal? (hash-ref (STYLES-font_style->index_map (*STYLES*)) "20<p>Arial<p>0000FF") 1)

       (check-equal? (hash-count (STYLES-number_style->index_map (*STYLES*))) 2)
       (check-equal? (hash-ref (STYLES-number_style->index_map (*STYLES*)) "0.00") 0)
       (check-equal? (hash-ref (STYLES-number_style->index_map (*STYLES*)) "0.00%") 1)

       (check-equal? (hash-count (STYLES-fill_style->index_map (*STYLES*))) 2)
       (check-equal? (hash-ref (STYLES-fill_style->index_map (*STYLES*)) "FF0000<p>solid") 0)
       (check-equal? (hash-ref (STYLES-fill_style->index_map (*STYLES*)) "FFF000<p>solid") 1)
       )))

   (test-case
    "test-sort-styles-cross-sheets1"

    (with-xlsx
     (lambda ()
       (add-data-sheet "Sheet1"
                       '(("month1" "month2" "month3" "month4" "real") (201601 100 110 1110 6.9)))
       (add-data-sheet "Sheet2" '((1)))
       (add-data-sheet "Sheet3" '((1)))
       (add-chart-sheet "Chart1" 'LINE "Chart1" '())
       (add-chart-sheet "Chart2" 'LINE "Chart2" '())
       (add-chart-sheet "Chart3" 'LINE "Chart3" '())

       (check-equal? (hash-count (STYLES-style->index_map (*STYLES*))) 0)

       (with-sheet-ref
        0
        (lambda ()
          (set-cell-range-border-style "A1-B2" "all" "0000ff" "dashed")
          (set-cell-range-border-style "B1-C2" "all" "0000ff" "thin")
          (set-cell-range-font-style "B1" 10 "Arial" "0000ff")
          (set-cell-range-font-style "B2" 20 "Arial" "0000ff")
          (set-cell-range-alignment-style "C1" "center" "center")
          (set-cell-range-alignment-style "C2" "center" "center")
          (set-cell-range-border-style "A1" "all" "0000f0" "dashed")
          (set-cell-range-number-style "A2" "0.001")))

       (with-sheet-ref
        1
        (lambda ()
          (set-cell-range-number-style "D1" "0.000")
          (set-cell-range-fill-style "D2" "FFF000" "solid")
          (set-cell-range-fill-style "E1" "FF0000" "solid")))

       (sort-styles)

       (check-equal? (hash-count (STYLES-style->index_map (*STYLES*))) 8)
       (check-equal? (hash-ref (STYLES-style->index_map (*STYLES*)) "<s><s><s><s>FF0000<p>solid") 0)
       (check-equal? (hash-ref (STYLES-style->index_map (*STYLES*)) "<s><s><s><s>FFF000<p>solid") 1)
       (check-equal? (hash-ref (STYLES-style->index_map (*STYLES*)) "<s><s><s>0.000<s>") 2)
       (check-equal? (hash-ref (STYLES-style->index_map (*STYLES*))
                               "0000F0<p>dashed<p>0000F0<p>dashed<p>0000F0<p>dashed<p>0000F0<p>dashed<s><s><s><s>") 3)
       (check-equal? (hash-ref (STYLES-style->index_map (*STYLES*))
                               "0000FF<p>dashed<p>0000FF<p>dashed<p>0000FF<p>dashed<p>0000FF<p>dashed<s><s><s>0.001<s>") 4)
       (check-equal? (hash-ref (STYLES-style->index_map (*STYLES*))
                               "0000FF<p>thin<p>0000FF<p>thin<p>0000FF<p>thin<p>0000FF<p>thin<s><s>center<p>center<s><s>") 5)
       (check-equal? (hash-ref (STYLES-style->index_map (*STYLES*))
                               "0000FF<p>thin<p>0000FF<p>thin<p>0000FF<p>thin<p>0000FF<p>thin<s>10<p>Arial<p>0000FF<s><s><s>") 6)
       (check-equal? (hash-ref (STYLES-style->index_map (*STYLES*))
                               "0000FF<p>thin<p>0000FF<p>thin<p>0000FF<p>thin<p>0000FF<p>thin<s>20<p>Arial<p>0000FF<s><s><s>") 7)

       (check-equal? (hash-count (STYLES-border_style->index_map (*STYLES*))) 3)
       (check-equal? (hash-ref (STYLES-border_style->index_map (*STYLES*))
                               "0000F0<p>dashed<p>0000F0<p>dashed<p>0000F0<p>dashed<p>0000F0<p>dashed") 0)
       (check-equal? (hash-ref (STYLES-border_style->index_map (*STYLES*))
                               "0000FF<p>dashed<p>0000FF<p>dashed<p>0000FF<p>dashed<p>0000FF<p>dashed") 1)
       (check-equal? (hash-ref (STYLES-border_style->index_map (*STYLES*))
                               "0000FF<p>thin<p>0000FF<p>thin<p>0000FF<p>thin<p>0000FF<p>thin") 2)

       (check-equal? (hash-count (STYLES-font_style->index_map (*STYLES*))) 2)
       (check-equal? (hash-ref (STYLES-font_style->index_map (*STYLES*)) "10<p>Arial<p>0000FF") 0)
       (check-equal? (hash-ref (STYLES-font_style->index_map (*STYLES*)) "20<p>Arial<p>0000FF") 1)

       (check-equal? (hash-count (STYLES-number_style->index_map (*STYLES*))) 2)
       (check-equal? (hash-ref (STYLES-number_style->index_map (*STYLES*)) "0.000") 0)
       (check-equal? (hash-ref (STYLES-number_style->index_map (*STYLES*)) "0.001") 1)

       (check-equal? (hash-count (STYLES-fill_style->index_map (*STYLES*))) 2)
       (check-equal? (hash-ref (STYLES-fill_style->index_map (*STYLES*)) "FF0000<p>solid") 0)
       (check-equal? (hash-ref (STYLES-fill_style->index_map (*STYLES*)) "FFF000<p>solid") 1)
       )))

   (test-case
    "test-sort-styles-cross-sheets2"

    (with-xlsx
     (lambda ()
       (add-data-sheet "Sheet1"
                       '(("month1" "month2" "month3" "month4" "real") (201601 100 110 1110 6.9)))
       (add-data-sheet "Sheet2" '((1)))
       (add-data-sheet "Sheet3" '((1)))
       (add-chart-sheet "Chart1" 'LINE "Chart1" '())
       (add-chart-sheet "Chart2" 'LINE "Chart2" '())
       (add-chart-sheet "Chart3" 'LINE "Chart3" '())

       (check-equal? (hash-count (STYLES-style->index_map (*STYLES*))) 0)

       (with-sheet-ref 0  (lambda () (set-cell-range-border-style "A1-B2" "all" "0000ff" "dashed")))
       (with-sheet-ref 1  (lambda () (set-cell-range-border-style "B1-C2" "all" "0000ff" "thin")))
       (with-sheet-ref 2  (lambda () (set-cell-range-font-style "B1" 10 "Arial" "0000ff")))

       (sort-styles)

       (check-equal? (hash-count (STYLES-style->index_map (*STYLES*))) 3)
       (check-equal? (hash-ref (STYLES-style->index_map (*STYLES*))
                               "<s>10<p>Arial<p>0000FF<s><s><s>") 0)
       (check-equal? (hash-ref (STYLES-style->index_map (*STYLES*))
                               "0000FF<p>dashed<p>0000FF<p>dashed<p>0000FF<p>dashed<p>0000FF<p>dashed<s><s><s><s>") 1)
       (check-equal? (hash-ref (STYLES-style->index_map (*STYLES*))
                               "0000FF<p>thin<p>0000FF<p>thin<p>0000FF<p>thin<p>0000FF<p>thin<s><s><s><s>") 2)

       (check-equal? (hash-count (STYLES-border_style->index_map (*STYLES*))) 2)
       (check-equal? (hash-ref (STYLES-border_style->index_map (*STYLES*)) "0000FF<p>dashed<p>0000FF<p>dashed<p>0000FF<p>dashed<p>0000FF<p>dashed") 0)
       (check-equal? (hash-ref (STYLES-border_style->index_map (*STYLES*)) "0000FF<p>thin<p>0000FF<p>thin<p>0000FF<p>thin<p>0000FF<p>thin") 1)

       (check-equal? (hash-count (STYLES-font_style->index_map (*STYLES*))) 1)
       (check-equal? (hash-ref (STYLES-font_style->index_map (*STYLES*)) "10<p>Arial<p>0000FF") 0)

       (check-equal? (hash-count (STYLES-number_style->index_map (*STYLES*))) 0)
       (check-equal? (hash-count (STYLES-fill_style->index_map (*STYLES*))) 0)
       )))
   ))

(run-tests test-styles)
