#lang racket

(require rackunit/text-ui
         rackunit
         "../../../xlsx/xlsx.rkt"
         "../../../sheet/sheet.rkt"
         "../../../style/style.rkt"
         "../../../style/styles.rkt"
         "../../../style/font-style.rkt"
         "../../../style/fill-style.rkt"
         "../../../style/alignment-style.rkt"
         "../../../style/border-style.rkt"
         "../../../style/number-style.rkt"
         "../../../style/assemble-styles.rkt"
         "../../../style/set-styles.rkt")

(define test-styles
  (test-suite
   "test-styles"

   (test-case
    "test-assemble-styles-normal"

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
          (check-equal? (length (STYLES-styles (*STYLES*))) 0)
          (check-equal? (hash-count (SHEET-STYLE-cell->style_map (*CURRENT_SHEET_STYLE*))) 0)
          (check-equal? (hash-count (SHEET-STYLE-row->style_map (*CURRENT_SHEET_STYLE*))) 0)
          (check-equal? (hash-count (SHEET-STYLE-col->style_map (*CURRENT_SHEET_STYLE*))) 0)

          (set-cell-range-border-style "A1" "top" "F00000" "thin")
          (set-cell-range-border-style "A1" "bottom" "0F0000" "thick")
          (set-cell-range-border-style "A1" "left" "00F000" "double")
          (set-cell-range-border-style "A1" "right" "000F00" "dashed")
          (set-cell-range-font-style "B1" 10 "Arial" "0000FF")
          (set-cell-range-alignment-style "C1" "center" "center")
          (set-cell-range-font-style "D1" 10 "Arial" "0000FF")
          (set-cell-range-number-style "D1" "0.000")
          (set-cell-range-fill-style "E1" "FF0000" "solid")))

       (strip-styles)
       (assemble-styles)

       (check-equal? (length (STYLES-styles (*STYLES*))) 7)

       (check-equal? (list-ref (STYLES-styles (*STYLES*)) 0)
                     (STYLE #f #f #f #f (FILL-STYLE "FFFFFF" "none")))
       (check-equal? (list-ref (STYLES-styles (*STYLES*)) 1)
                     (STYLE #f #f #f #f (FILL-STYLE "FFFFFF" "gray125")))
       (check-equal? (list-ref (STYLES-styles (*STYLES*)) 2)
                     (STYLE
                      (BORDER-STYLE "00F000" "double" "000F00" "dashed" "F00000" "thin" "0F0000" "thick")
                      #f #f #f #f))
       (check-equal? (list-ref (STYLES-styles (*STYLES*)) 3) (STYLE #f (FONT-STYLE 10 "Arial" "0000FF") #f #f #f))
       (check-equal? (list-ref (STYLES-styles (*STYLES*)) 4)
                     (STYLE #f #f (ALIGNMENT-STYLE "center" "center") #f #f))
       (check-equal? (list-ref (STYLES-styles (*STYLES*)) 5)
                     (STYLE #f (FONT-STYLE 10 "Arial" "0000FF") #f (NUMBER-STYLE "1" "0.000") #f))
       (check-equal? (list-ref (STYLES-styles (*STYLES*)) 6)
                     (STYLE #f #f #f #f (FILL-STYLE "FF0000" "solid")))

       (with-sheet-ref
        0
        (lambda ()
          (check-equal? (hash-count (SHEET-STYLE-cell->style_map (*CURRENT_SHEET_STYLE*))) 5)
          (check-equal? (hash-ref (SHEET-STYLE-cell->style_map (*CURRENT_SHEET_STYLE*)) "A1")
                        (STYLE
                         (BORDER-STYLE "00F000" "double" "000F00" "dashed" "F00000" "thin" "0F0000" "thick")
                         #f #f #f #f))
          (check-equal? (hash-ref (SHEET-STYLE-cell->style_map (*CURRENT_SHEET_STYLE*)) "B1")
                        (STYLE
                         #f (FONT-STYLE 10 "Arial" "0000FF") #f #f #f))
          (check-equal? (hash-ref (SHEET-STYLE-cell->style_map (*CURRENT_SHEET_STYLE*)) "C1")
                        (STYLE
                         #f #f (ALIGNMENT-STYLE "center" "center") #f #f))
          (check-equal? (hash-ref (SHEET-STYLE-cell->style_map (*CURRENT_SHEET_STYLE*)) "D1")
                        (STYLE
                         #f (FONT-STYLE 10 "Arial" "0000FF") #f (NUMBER-STYLE "1" "0.000") #f))
          (check-equal? (hash-ref (SHEET-STYLE-cell->style_map (*CURRENT_SHEET_STYLE*)) "E1")
                        (STYLE
                         #f #f #f #f (FILL-STYLE "FF0000" "solid")))))

       (check-equal? (length (STYLES-border_list (*STYLES*))) 2)
       (check-equal? (list-ref (STYLES-border_list (*STYLES*)) 0)
                     (BORDER-STYLE #f #f #f #f #f #f #f #f))
       (check-equal? (list-ref
                      (STYLES-border_list (*STYLES*))
                      1)
                     (BORDER-STYLE "00F000" "double" "000F00" "dashed" "F00000" "thin" "0F0000" "thick"))

       (check-equal? (length (STYLES-font_list (*STYLES*))) 2)
       (check-equal? (list-ref (STYLES-font_list (*STYLES*)) 0)
                     (FONT-STYLE 10 "Arial" "000000"))
       (check-equal? (list-ref (STYLES-font_list (*STYLES*)) 1)
                     (FONT-STYLE 10 "Arial" "0000FF"))

       (check-equal? (length (STYLES-number_list (*STYLES*))) 1)
       (check-equal? (list-ref (STYLES-number_list (*STYLES*)) 0)
                     (NUMBER-STYLE "1" "0.000"))

       (check-equal? (length (STYLES-fill_list (*STYLES*))) 3)
       (check-equal? (list-ref (STYLES-fill_list (*STYLES*)) 0)
                     (FILL-STYLE "FFFFFF" "none"))
       (check-equal? (list-ref (STYLES-fill_list (*STYLES*)) 1)
                     (FILL-STYLE "FFFFFF" "gray125"))
       (check-equal? (list-ref (STYLES-fill_list (*STYLES*)) 2)
                     (FILL-STYLE "FF0000" "solid"))
       )))

   (test-case
    "test-assemble-styles-from-read"

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
          (check-equal? (length (STYLES-styles (*STYLES*))) 0)
          (check-equal? (hash-count (SHEET-STYLE-cell->style_map (*CURRENT_SHEET_STYLE*))) 0)
          (check-equal? (hash-count (SHEET-STYLE-row->style_map (*CURRENT_SHEET_STYLE*))) 0)
          (check-equal? (hash-count (SHEET-STYLE-col->style_map (*CURRENT_SHEET_STYLE*))) 0)

          (set-cell-range-border-style "A1" "top" "F00000" "thin")
          (set-cell-range-border-style "A1" "bottom" "0F0000" "thick")
          (set-cell-range-border-style "A1" "left" "00F000" "double")
          (set-cell-range-border-style "A1" "right" "000F00" "dashed")
          (set-cell-range-font-style "B1" 10 "Arial" "0000FF")
          (set-cell-range-alignment-style "C1" "center" "center")
          (set-cell-range-font-style "D1" 10 "Arial" "0000FF")
          (set-cell-range-number-style "D1" "0.000")
          (set-cell-range-fill-style "E1" "FF0000" "solid")))

       (strip-styles)

       (check-equal? (length (STYLES-styles (*STYLES*))) 5)
       (check-equal? (list-ref (STYLES-styles (*STYLES*)) 0)
                     (STYLE
                      (BORDER-STYLE "00F000" "double" "000F00" "dashed" "F00000" "thin" "0F0000" "thick")
                      #f #f #f #f))
       (check-equal? (list-ref (STYLES-styles (*STYLES*)) 1) (STYLE #f (FONT-STYLE 10 "Arial" "0000FF") #f #f #f))
       (check-equal? (list-ref (STYLES-styles (*STYLES*)) 2)
                     (STYLE #f #f (ALIGNMENT-STYLE "center" "center") #f #f))
       (check-equal? (list-ref (STYLES-styles (*STYLES*)) 3)
                     (STYLE #f (FONT-STYLE 10 "Arial" "0000FF") #f (NUMBER-STYLE "1" "0.000") #f))
       (check-equal? (list-ref (STYLES-styles (*STYLES*)) 4)
                     (STYLE #f #f #f #f (FILL-STYLE "FF0000" "solid")))

       (with-sheet-ref
        0
        (lambda ()
          (check-equal? (hash-count (SHEET-STYLE-cell->style_map (*CURRENT_SHEET_STYLE*))) 5)
          (check-equal? (hash-ref (SHEET-STYLE-cell->style_map (*CURRENT_SHEET_STYLE*)) "A1")
                        (STYLE
                         (BORDER-STYLE "00F000" "double" "000F00" "dashed" "F00000" "thin" "0F0000" "thick")
                         #f #f #f #f))
          (check-equal? (hash-ref (SHEET-STYLE-cell->style_map (*CURRENT_SHEET_STYLE*)) "B1")
                        (STYLE
                         #f (FONT-STYLE 10 "Arial" "0000FF") #f #f #f))
          (check-equal? (hash-ref (SHEET-STYLE-cell->style_map (*CURRENT_SHEET_STYLE*)) "C1")
                        (STYLE
                         #f #f (ALIGNMENT-STYLE "center" "center") #f #f))
          (check-equal? (hash-ref (SHEET-STYLE-cell->style_map (*CURRENT_SHEET_STYLE*)) "D1")
                        (STYLE
                         #f (FONT-STYLE 10 "Arial" "0000FF") #f (NUMBER-STYLE "1" "0.000") #f))
          (check-equal? (hash-ref (SHEET-STYLE-cell->style_map (*CURRENT_SHEET_STYLE*)) "E1")
                        (STYLE
                         #f #f #f #f (FILL-STYLE "FF0000" "solid")))))

       (check-equal? (length (STYLES-border_list (*STYLES*))) 1)
       (check-equal? (list-ref
                      (STYLES-border_list (*STYLES*))
                      0)
                     (BORDER-STYLE "00F000" "double" "000F00" "dashed" "F00000" "thin" "0F0000" "thick"))

       (check-equal? (length (STYLES-font_list (*STYLES*))) 1)
       (check-equal? (list-ref (STYLES-font_list (*STYLES*)) 0)
                     (FONT-STYLE 10 "Arial" "0000FF"))

       (check-equal? (length (STYLES-number_list (*STYLES*))) 1)
       (check-equal? (list-ref (STYLES-number_list (*STYLES*)) 0)
                     (NUMBER-STYLE "1" "0.000"))

       (check-equal? (length (STYLES-fill_list (*STYLES*))) 1)
       (check-equal? (list-ref (STYLES-fill_list (*STYLES*)) 0)
                     (FILL-STYLE "FF0000" "solid"))
       )))

   (test-case
    "test-assemble-styles-overlap"

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
          (check-equal? (length (STYLES-styles (*STYLES*))) 0)

          (set-cell-range-border-style "A1" "top" "F00000" "thin")
          (set-cell-range-border-style "A1" "bottom" "0F0000" "thick")
          (set-cell-range-border-style "A1" "left" "00F000" "double")
          (set-cell-range-border-style "A1" "right" "000F00" "dashed")

          (set-cell-range-font-style "B1" 10 "Arial" "0000FF")
          (set-cell-range-alignment-style "C1" "center" "center")
          (set-cell-range-number-style "D1" "0.000")
          (set-cell-range-fill-style "E1" "FF0000" "solid")

          (set-cell-range-border-style "A1" "top" "F00000" "thin")
          (set-cell-range-border-style "A1" "bottom" "0F0000" "thick")
          (set-cell-range-border-style "A1" "left" "00F000" "double")
          (set-cell-range-border-style "A1" "right" "00000F" "dashed")

          (strip-styles)
          (assemble-styles)

          (check-equal? (length (STYLES-styles (*STYLES*))) 7)
          (check-equal?
           (list-ref (STYLES-styles (*STYLES*)) 6)
            (STYLE
             (BORDER-STYLE "00F000" "double" "00000F" "dashed" "F00000" "thin" "0F0000" "thick") #f #f #f #f))
          )))))

   (test-case
    "test-assemble-styles-overlap-partly"

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
          (check-equal? (length (STYLES-styles (*STYLES*))) 0)

          (set-cell-range-border-style "A1-B2" "all" "0000FF" "dashed")
          (set-cell-range-border-style "B1-C2" "all" "0000FF" "thin")
          (set-cell-range-font-style "B1" 10 "Arial" "0000FF")
          (set-cell-range-font-style "B2" 20 "Arial" "0000FF")
          (set-cell-range-alignment-style "C1" "center" "center")
          (set-cell-range-alignment-style "C2" "center" "center")
          (set-cell-range-number-style "D1" "0.00")
          (set-cell-range-number-style "A2" "0.00%")
          (set-cell-range-fill-style "E1" "FF0000" "solid")
          (set-cell-range-fill-style "D2" "FFF000" "solid")
          (set-cell-range-border-style "A1" "all" "0000F0" "dashed")
          ))

       (strip-styles)
       (assemble-styles)

       (check-equal? (length (STYLES-styles (*STYLES*))) 10)
       (check-equal? (list-ref (STYLES-styles (*STYLES*)) 0)
                     (STYLE #f #f #f #f (FILL-STYLE "FFFFFF" "none")))
       (check-equal? (list-ref (STYLES-styles (*STYLES*)) 1)
                     (STYLE #f #f #f #f (FILL-STYLE "FFFFFF" "gray125")))
       (check-equal? (list-ref (STYLES-styles (*STYLES*)) 2)
                     (STYLE
                      (BORDER-STYLE "0000FF" "thin" "0000FF" "thin" "0000FF" "thin" "0000FF" "thin") (FONT-STYLE 10 "Arial" "0000FF") #f #f #f))
       (check-equal? (list-ref (STYLES-styles (*STYLES*)) 3)
                     (STYLE
                      (BORDER-STYLE "0000FF" "thin" "0000FF" "thin" "0000FF" "thin" "0000FF" "thin") (FONT-STYLE 20 "Arial" "0000FF") #f #f #f))
       (check-equal? (list-ref (STYLES-styles (*STYLES*)) 4)
                     (STYLE
                      (BORDER-STYLE "0000FF" "thin" "0000FF" "thin" "0000FF" "thin" "0000FF" "thin") #f (ALIGNMENT-STYLE "center" "center") #f #f))
       (check-equal? (list-ref (STYLES-styles (*STYLES*)) 5)
                     (STYLE #f #f #f (NUMBER-STYLE "1" "0.00") #f))
       (check-equal? (list-ref (STYLES-styles (*STYLES*)) 6)
                     (STYLE
                      (BORDER-STYLE "0000FF" "dashed" "0000FF" "dashed" "0000FF" "dashed" "0000FF" "dashed") #f #f (NUMBER-STYLE "2" "0.00%") #f))
       (check-equal? (list-ref (STYLES-styles (*STYLES*)) 7)
                     (STYLE #f #f #f #f (FILL-STYLE "FF0000" "solid")))
       (check-equal? (list-ref (STYLES-styles (*STYLES*)) 8)
                     (STYLE #f #f #f #f (FILL-STYLE "FFF000" "solid")))
       (check-equal? (list-ref (STYLES-styles (*STYLES*)) 9)
                     (STYLE
                      (BORDER-STYLE "0000F0" "dashed" "0000F0" "dashed" "0000F0" "dashed" "0000F0" "dashed")
                      #f #f #f #f))

       (with-sheet-ref
        0
        (lambda ()
          (check-equal? (hash-count (SHEET-STYLE-cell->style_map (*CURRENT_SHEET_STYLE*))) 9)
          (check-equal? (hash-ref (SHEET-STYLE-cell->style_map (*CURRENT_SHEET_STYLE*)) "A1")
                        (STYLE
                         (BORDER-STYLE "0000F0" "dashed" "0000F0" "dashed" "0000F0" "dashed" "0000F0" "dashed")
                         #f #f #f #f))
          (check-equal? (hash-ref (SHEET-STYLE-cell->style_map (*CURRENT_SHEET_STYLE*)) "A2")
                        (STYLE
                         (BORDER-STYLE "0000FF" "dashed" "0000FF" "dashed" "0000FF" "dashed" "0000FF" "dashed")
                         #f #f (NUMBER-STYLE "2" "0.00%") #f))
          (check-equal? (hash-ref (SHEET-STYLE-cell->style_map (*CURRENT_SHEET_STYLE*)) "B1")
                        (STYLE
                         (BORDER-STYLE "0000FF" "thin" "0000FF" "thin" "0000FF" "thin" "0000FF" "thin")
                         (FONT-STYLE 10 "Arial" "0000FF") #f #f #f))
          (check-equal? (hash-ref (SHEET-STYLE-cell->style_map (*CURRENT_SHEET_STYLE*)) "B2")
                        (STYLE
                         (BORDER-STYLE "0000FF" "thin" "0000FF" "thin" "0000FF" "thin" "0000FF" "thin")
                         (FONT-STYLE 20 "Arial" "0000FF") #f #f #f))
          (check-equal? (hash-ref (SHEET-STYLE-cell->style_map (*CURRENT_SHEET_STYLE*)) "C1")
                        (STYLE
                         (BORDER-STYLE "0000FF" "thin" "0000FF" "thin" "0000FF" "thin" "0000FF" "thin")
                         #f (ALIGNMENT-STYLE "center" "center") #f #f))
          (check-equal? (hash-ref (SHEET-STYLE-cell->style_map (*CURRENT_SHEET_STYLE*)) "C2")
                        (STYLE
                         (BORDER-STYLE "0000FF" "thin" "0000FF" "thin" "0000FF" "thin" "0000FF" "thin")
                         #f (ALIGNMENT-STYLE "center" "center") #f #f))
          (check-equal? (hash-ref (SHEET-STYLE-cell->style_map (*CURRENT_SHEET_STYLE*)) "D1")
                        (STYLE #f #f #f (NUMBER-STYLE "1" "0.00") #f))
          (check-equal? (hash-ref (SHEET-STYLE-cell->style_map (*CURRENT_SHEET_STYLE*)) "D2")
                        (STYLE #f #f #f #f (FILL-STYLE "FFF000" "solid")))
          (check-equal? (hash-ref (SHEET-STYLE-cell->style_map (*CURRENT_SHEET_STYLE*)) "E1")
                        (STYLE #f #f #f #f (FILL-STYLE "FF0000" "solid")))
          ))

       (check-equal? (length (STYLES-border_list (*STYLES*))) 4)
       (check-equal?
        (list-ref (STYLES-border_list (*STYLES*)) 0)
         (BORDER-STYLE #f #f #f #f #f #f #f #f))
       (check-equal?
        (list-ref (STYLES-border_list (*STYLES*)) 1)
         (BORDER-STYLE "0000FF" "dashed" "0000FF" "dashed" "0000FF" "dashed" "0000FF" "dashed"))
       (check-equal?
        (list-ref (STYLES-border_list (*STYLES*)) 2)
        (BORDER-STYLE "0000FF" "thin" "0000FF" "thin" "0000FF" "thin" "0000FF" "thin"))
       (check-equal?
        (list-ref (STYLES-border_list (*STYLES*)) 3)
        (BORDER-STYLE "0000F0" "dashed" "0000F0" "dashed" "0000F0" "dashed" "0000F0" "dashed"))

       (check-equal? (length (STYLES-font_list (*STYLES*))) 3)
       (check-equal? (list-ref (STYLES-font_list (*STYLES*)) 0) (FONT-STYLE 10 "Arial" "000000"))
       (check-equal? (list-ref (STYLES-font_list (*STYLES*)) 1) (FONT-STYLE 10 "Arial" "0000FF"))
       (check-equal? (list-ref (STYLES-font_list (*STYLES*)) 2) (FONT-STYLE 20 "Arial" "0000FF"))

       (check-equal? (length (STYLES-number_list (*STYLES*))) 2)
       (check-equal? (list-ref (STYLES-number_list (*STYLES*)) 0) (NUMBER-STYLE "1" "0.00"))
       (check-equal? (list-ref (STYLES-number_list (*STYLES*)) 1) (NUMBER-STYLE "2" "0.00%"))

       (check-equal? (length (STYLES-fill_list (*STYLES*))) 4)
       (check-equal? (list-ref (STYLES-fill_list (*STYLES*)) 0) (FILL-STYLE "FFFFFF" "none"))
       (check-equal? (list-ref (STYLES-fill_list (*STYLES*)) 1) (FILL-STYLE "FFFFFF" "gray125"))
       (check-equal? (list-ref (STYLES-fill_list (*STYLES*)) 2) (FILL-STYLE "FF0000" "solid"))
       (check-equal? (list-ref (STYLES-fill_list (*STYLES*)) 3) (FILL-STYLE "FFF000" "solid"))
       )))

   (test-case
    "test-assemble-styles-cross-sheets1"

    (with-xlsx
     (lambda ()
       (add-data-sheet "Sheet1"
                       '(("month1" "month2" "month3" "month4" "real") (201601 100 110 1110 6.9)))
       (add-data-sheet "Sheet2" '((1)))
       (add-data-sheet "Sheet3" '((1)))
       (add-chart-sheet "Chart1" 'LINE "Chart1" '())
       (add-chart-sheet "Chart2" 'LINE "Chart2" '())
       (add-chart-sheet "Chart3" 'LINE "Chart3" '())

       (check-equal? (length (STYLES-styles (*STYLES*))) 0)

       (with-sheet-ref
        0
        (lambda ()
          (set-cell-range-border-style "A1-B2" "all" "0000FF" "dashed")
          (set-cell-range-border-style "B1-C2" "all" "0000FF" "thin")
          (set-cell-range-font-style "B1" 10 "Arial" "0000FF")
          (set-cell-range-font-style "B2" 20 "Arial" "0000FF")
          (set-cell-range-alignment-style "C1" "center" "center")
          (set-cell-range-alignment-style "C2" "center" "center")
          (set-cell-range-border-style "A1" "all" "0000F0" "dashed")
          (set-cell-range-number-style "A2" "0.001")))

       (with-sheet-ref
        1
        (lambda ()
          (set-cell-range-number-style "D1" "0.000")
          (set-cell-range-fill-style "D2" "FFF000" "solid")
          (set-cell-range-fill-style "E1" "FF0000" "solid")))

       (strip-styles)
       (assemble-styles)

       (check-equal? (length (STYLES-styles (*STYLES*))) 10)
       (check-equal? (list-ref (STYLES-styles (*STYLES*)) 0)
                     (STYLE #f #f #f #f (FILL-STYLE "FFFFFF" "none")))
       (check-equal? (list-ref (STYLES-styles (*STYLES*)) 1)
                     (STYLE #f #f #f #f (FILL-STYLE "FFFFFF" "gray125")))
       (check-equal? (list-ref (STYLES-styles (*STYLES*)) 2)
                     (STYLE
                      (BORDER-STYLE "0000FF" "thin" "0000FF" "thin" "0000FF" "thin" "0000FF" "thin") (FONT-STYLE 10 "Arial" "0000FF") #f #f #f))
       (check-equal? (list-ref (STYLES-styles (*STYLES*)) 3)
                     (STYLE
                      (BORDER-STYLE "0000FF" "thin" "0000FF" "thin" "0000FF" "thin" "0000FF" "thin") (FONT-STYLE 20 "Arial" "0000FF") #f #f #f))
       (check-equal? (list-ref (STYLES-styles (*STYLES*)) 4)
                     (STYLE
                      (BORDER-STYLE "0000FF" "thin" "0000FF" "thin" "0000FF" "thin" "0000FF" "thin") #f (ALIGNMENT-STYLE "center" "center") #f #f))
       (check-equal? (list-ref (STYLES-styles (*STYLES*)) 5)
                     (STYLE
                      (BORDER-STYLE "0000F0" "dashed" "0000F0" "dashed" "0000F0" "dashed" "0000F0" "dashed")
                      #f #f #f #f))
       (check-equal? (list-ref (STYLES-styles (*STYLES*)) 6)
                     (STYLE
                      (BORDER-STYLE "0000FF" "dashed" "0000FF" "dashed" "0000FF" "dashed" "0000FF" "dashed") #f #f (NUMBER-STYLE "1" "0.001") #f))
       (check-equal? (list-ref (STYLES-styles (*STYLES*)) 7)
                     (STYLE #f #f #f (NUMBER-STYLE "2" "0.000") #f))
       (check-equal? (list-ref (STYLES-styles (*STYLES*)) 8)
                     (STYLE #f #f #f #f (FILL-STYLE "FFF000" "solid")))
       (check-equal? (list-ref (STYLES-styles (*STYLES*)) 9)
                     (STYLE #f #f #f #f (FILL-STYLE "FF0000" "solid")))

       (check-equal? (length (STYLES-border_list (*STYLES*))) 4)
       (check-equal? (list-ref (STYLES-border_list (*STYLES*)) 0)
                     (BORDER-STYLE #f #f #f #f #f #f #f #f))
       (check-equal? (list-ref (STYLES-border_list (*STYLES*)) 1)
                     (BORDER-STYLE "0000FF" "dashed" "0000FF" "dashed" "0000FF" "dashed" "0000FF" "dashed"))
       (check-equal? (list-ref (STYLES-border_list (*STYLES*)) 2)
                     (BORDER-STYLE "0000FF" "thin" "0000FF" "thin" "0000FF" "thin" "0000FF" "thin"))
       (check-equal? (list-ref (STYLES-border_list (*STYLES*)) 3)
                     (BORDER-STYLE "0000F0" "dashed" "0000F0" "dashed" "0000F0" "dashed" "0000F0" "dashed"))

       (check-equal? (length (STYLES-font_list (*STYLES*))) 3)
       (check-equal? (list-ref (STYLES-font_list (*STYLES*)) 0) (FONT-STYLE 10 "Arial" "000000"))
       (check-equal? (list-ref (STYLES-font_list (*STYLES*)) 1) (FONT-STYLE 10 "Arial" "0000FF"))
       (check-equal? (list-ref (STYLES-font_list (*STYLES*)) 2) (FONT-STYLE 20 "Arial" "0000FF"))

       (check-equal? (length (STYLES-number_list (*STYLES*))) 2)
       (check-equal? (list-ref (STYLES-number_list (*STYLES*)) 0) (NUMBER-STYLE "1" "0.001"))
       (check-equal? (list-ref (STYLES-number_list (*STYLES*)) 1) (NUMBER-STYLE "2" "0.000"))

       (check-equal? (length (STYLES-fill_list (*STYLES*))) 4)
       (check-equal? (list-ref (STYLES-fill_list (*STYLES*)) 0) (FILL-STYLE "FFFFFF" "none"))
       (check-equal? (list-ref (STYLES-fill_list (*STYLES*)) 1) (FILL-STYLE "FFFFFF" "gray125"))
       (check-equal? (list-ref (STYLES-fill_list (*STYLES*)) 2) (FILL-STYLE "FFF000" "solid"))
       (check-equal? (list-ref (STYLES-fill_list (*STYLES*)) 3) (FILL-STYLE "FF0000" "solid"))
       )))

   (test-case
    "test-assemble-styles-cross-sheets2"

    (with-xlsx
     (lambda ()
       (add-data-sheet "Sheet1"
                       '(("month1" "month2" "month3" "month4" "real") (201601 100 110 1110 6.9)))
       (add-data-sheet "Sheet2" '((1)))
       (add-data-sheet "Sheet3" '((1)))
       (add-chart-sheet "Chart1" 'LINE "Chart1" '())
       (add-chart-sheet "Chart2" 'LINE "Chart2" '())
       (add-chart-sheet "Chart3" 'LINE "Chart3" '())

       (check-equal? (length (STYLES-styles (*STYLES*))) 0)

       (with-sheet-ref 0  (lambda () (set-cell-range-border-style "A1-B2" "all" "0000FF" "dashed")))
       (with-sheet-ref 1  (lambda () (set-cell-range-border-style "B1-C2" "all" "0000FF" "thin")))
       (with-sheet-ref 2  (lambda () (set-cell-range-font-style "B1" 10 "Arial" "0000FF")))

       (strip-styles)
       (assemble-styles)

       (check-equal? (length (STYLES-styles (*STYLES*))) 5)
       (check-equal? (list-ref (STYLES-styles (*STYLES*)) 0)
                     (STYLE #f #f #f #f (FILL-STYLE "FFFFFF" "none")))
       (check-equal? (list-ref (STYLES-styles (*STYLES*)) 1)
                     (STYLE #f #f #f #f (FILL-STYLE "FFFFFF" "gray125")))
       (check-equal? (list-ref (STYLES-styles (*STYLES*)) 2)
                     (STYLE
                      (BORDER-STYLE "0000FF" "dashed" "0000FF" "dashed" "0000FF" "dashed" "0000FF" "dashed")
                      #f #f #f #f))
       (check-equal? (list-ref (STYLES-styles (*STYLES*)) 3)
                     (STYLE
                      (BORDER-STYLE "0000FF" "thin" "0000FF" "thin" "0000FF" "thin" "0000FF" "thin")
                      #f #f #f #f))
       (check-equal? (list-ref (STYLES-styles (*STYLES*)) 4)
                     (STYLE
                      #f (FONT-STYLE 10 "Arial" "0000FF") #f #f #f))

       (check-equal? (length (STYLES-border_list (*STYLES*))) 3)
       (check-equal? (list-ref (STYLES-border_list (*STYLES*)) 0)
                     (BORDER-STYLE #f #f #f #f #f #f #f #f))
       (check-equal? (list-ref (STYLES-border_list (*STYLES*)) 1)
                     (BORDER-STYLE "0000FF" "dashed" "0000FF" "dashed" "0000FF" "dashed" "0000FF" "dashed"))
       (check-equal? (list-ref (STYLES-border_list (*STYLES*)) 2)
                     (BORDER-STYLE "0000FF" "thin" "0000FF" "thin" "0000FF" "thin" "0000FF" "thin"))

       (check-equal? (length (STYLES-font_list (*STYLES*))) 2)
       (check-equal? (list-ref (STYLES-font_list (*STYLES*)) 0) (FONT-STYLE 10 "Arial" "000000"))
       (check-equal? (list-ref (STYLES-font_list (*STYLES*)) 1) (FONT-STYLE 10 "Arial" "0000FF"))

       (check-equal? (length (STYLES-number_list (*STYLES*))) 0)
       (check-equal? (length (STYLES-fill_list (*STYLES*))) 2)
       )))

   (test-case
    "remove duplicate struct"

    (let ([fix_head_styles
           (list (FILL-STYLE "FFFFFF" "none")
                 (FILL-STYLE "FFFFFF" "gray125"))]
          [fill_styles
           (list (FILL-STYLE "FFFF00" "solid")
                 (FILL-STYLE "FFFFFF" "none")
                 (FILL-STYLE "FF0000" "solid")
                 (FILL-STYLE "FFFFFF" "gray125")
                 (FILL-STYLE "00FFFF" "lightDown")
                 )])
      (check-equal?
       (append
        fix_head_styles
        (remove*
         fix_head_styles
         fill_styles
         (lambda (a b) (= (equal-hash-code a) (equal-hash-code b)))))
       (list (FILL-STYLE "FFFFFF" "none")
             (FILL-STYLE "FFFFFF" "gray125")
             (FILL-STYLE "FFFF00" "solid")
             (FILL-STYLE "FF0000" "solid")
             (FILL-STYLE "00FFFF" "lightDown"))))

    (let ([fix_head_styles
           (list (STYLE #f #f (ALIGNMENT-STYLE "left" "center") #f (FILL-STYLE "FFFFFF" "none"))
                 (STYLE #f #f (ALIGNMENT-STYLE "left" "center") #f (FILL-STYLE "FFFFFF" "gray125")))]
          [styles
           (list
            (STYLE #f #f (ALIGNMENT-STYLE "left" "center") #f (FILL-STYLE "FFFFFF" "solid"))
            (STYLE #f #f (ALIGNMENT-STYLE "left" "center") #f (FILL-STYLE "FFFFFF" "none"))
            (STYLE #f #f (ALIGNMENT-STYLE "left" "center") #f (FILL-STYLE "FFFF00" "solid"))
            (STYLE #f #f (ALIGNMENT-STYLE "left" "center") #f (FILL-STYLE "FFFFFF" "gray125")))])

      (check-equal?
       (append
        fix_head_styles
        (remove*
         (list (STYLE #f #f (ALIGNMENT-STYLE "left" "center") #f (FILL-STYLE "FFFFFF" "none"))
               (STYLE #f #f (ALIGNMENT-STYLE "left" "center") #f (FILL-STYLE "FFFFFF" "gray125")))
         styles
         (lambda (a b) (= (equal-hash-code a) (equal-hash-code b)))))
       (list
        (STYLE #f #f (ALIGNMENT-STYLE "left" "center") #f (FILL-STYLE "FFFFFF" "none"))
        (STYLE #f #f (ALIGNMENT-STYLE "left" "center") #f (FILL-STYLE "FFFFFF" "gray125"))
        (STYLE #f #f (ALIGNMENT-STYLE "left" "center") #f (FILL-STYLE "FFFFFF" "solid"))
        (STYLE #f #f (ALIGNMENT-STYLE "left" "center") #f (FILL-STYLE "FFFF00" "solid")))))
    )

   ))

(run-tests test-styles)
