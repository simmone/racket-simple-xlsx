#lang racket

(require simple-xml)

(require rackunit/text-ui rackunit)

(require "../../../../xlsx/xlsx.rkt")
(require "../../../../sheet/sheet.rkt")
(require "../../../../lib/lib.rkt")
(require "../../../../xl/charts/lib.rkt")
(require"../../../../xl/charts/pie-chart.rkt")
(require"../../../../xl/charts/lib.rkt")

(require racket/runtime-path)
(define-runtime-path pie_chart_file "pie_chart.xml")

(define test-pie-chart
  (test-suite
   "test-pie-chart"

   (test-case
    "test-pie-chart"

    (with-xlsx
     (lambda ()
       (add-data-sheet "DataSheet"
                       '(
                         ("month1" "201601" "201602" "201603" "real")
                         (201601 100 300 200 6.9)
                         (201601 200 400 300 6.9)
                         (201601 300 500 400 6.9)
                         ))
       (add-data-sheet "Sheet2" '((1)))
       (add-data-sheet "Sheet3" '((1)))
       (add-chart-sheet "Chart1" 'PIE "Chart1" '())
       (add-chart-sheet "Chart2" 'PIE "Chart2" '())
       (add-chart-sheet "Chart3" 'PIE "Chart3" '())

       (call-with-input-file pie_chart_file
         (lambda (expected)
           (call-with-input-string
            (lists->xml_content
             (to-pie-chart-sers '(("CAT" "DataSheet" "B1-D1" "DataSheet" "B2-D2"))))
            (lambda (actual)
              (check-lines? expected actual))))))))
   ))

(run-tests test-pie-chart)
