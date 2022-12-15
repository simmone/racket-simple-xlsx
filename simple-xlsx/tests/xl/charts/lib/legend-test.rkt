#lang racket

(require simple-xml)

(require rackunit/text-ui rackunit)

(require "../../../../xlsx/xlsx.rkt")
(require "../../../../sheet/sheet.rkt")
(require "../../../../lib/lib.rkt")

(require"../../../../xl/charts/lib.rkt")

(require racket/runtime-path)
(define-runtime-path legend_file "legend.xml")

(define test-legend
  (test-suite
   "test-legend"

   (test-case
    "test-legend"

    (with-xlsx
     (lambda ()
      (add-data-sheet "DataSheet"
                      '(("month1" "201601" "201602" "201603" "real") (201601 100 300 200 6.9)))
      (add-data-sheet "Sheet2" '((1)))
      (add-data-sheet "Sheet3" '((1)))
      (add-chart-sheet "Chart1" 'LINE "Chart1" '())
      (add-chart-sheet "Chart2" 'LINE "Chart2" '())
      (add-chart-sheet "Chart3" 'LINE "Chart3" '())

      (with-sheet
       (lambda ()
         (call-with-input-file legend_file
         (lambda (expected)
           (call-with-input-string
            (lists->xml_content (legend))
            (lambda (actual)
              (check-lines? expected actual))))))))))
   ))

(run-tests test-legend)

