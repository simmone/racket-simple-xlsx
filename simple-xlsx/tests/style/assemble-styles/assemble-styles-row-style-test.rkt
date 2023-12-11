#lang racket

(require rackunit/text-ui rackunit)

(require "../../../xlsx/xlsx.rkt")
(require "../../../sheet/sheet.rkt")
(require "../../../style/style.rkt")
(require "../../../style/styles.rkt")
(require "../../../style/font-style.rkt")
(require "../../../style/fill-style.rkt")
(require "../../../style/assemble-styles.rkt")
(require "../../../style/set-styles.rkt")

(define test-styles
  (test-suite
   "test-styles"

   (test-case
    "test-assemble-styles"

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
          (check-equal? (hash-count (SHEET-STYLE-row->style_map (*CURRENT_SHEET_STYLE*))) 0)

          (set-row-range-font-style "1" 10 "Arial" "0000FF")
          ))

       (strip-styles)
       (assemble-styles)

       (check-equal? (length (STYLES-styles (*STYLES*))) 3)
       (check-equal? (list-ref (STYLES-styles (*STYLES*)) 0)
                     (STYLE
                      #f #f #f #f (FILL-STYLE "FFFFFF" "none")))
       (check-equal? (list-ref (STYLES-styles (*STYLES*)) 1)
                     (STYLE
                      #f #f #f #f (FILL-STYLE "FFFFFF" "gray125")))
       (check-equal? (list-ref (STYLES-styles (*STYLES*)) 2)
                     (STYLE
                      #f (FONT-STYLE 10 "Arial" "0000FF") #f #f #f))
       )))
   ))

(run-tests test-styles)
