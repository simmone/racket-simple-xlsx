#lang racket

(require fast-xml
         rackunit/text-ui
         rackunit
         "../../../../xlsx/xlsx.rkt"
         "../../../../sheet/sheet.rkt"
         "../../../../lib/lib.rkt"
         "../../../../xl/charts/line-chart.rkt"
         "../../../../xl/charts/charts-lib.rkt"
         racket/runtime-path)

(define-runtime-path line_3d_chart_file "line_3d_chart.xml")

(define test-line-3d-chart
  (test-suite
   "test-line-3d-chart"

   (test-case
    "test-line-3d-chart"

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
       (add-chart-sheet "Chart1" 'LINE "Chart1" '())
       (add-chart-sheet "Chart2" 'LINE "Chart2" '())
       (add-chart-sheet "Chart3" 'LINE "Chart3" '())

       (call-with-input-file line_3d_chart_file
         (lambda (expected)
           (call-with-input-string
            (lists-to-xml_content
             (to-line-3d-chart-sers
              '(
                ("CAT" "DataSheet" "B1-D1" "DataSheet" "B2-D2")
                ("Puma" "DataSheet" "B1-D1" "DataSheet" "B3-D3")
                ("Brooks" "DataSheet" "B1-D1" "DataSheet" "B4-D4")
                )))
            (lambda (actual)
              (check-lines? expected actual))))))))
   ))

(run-tests test-line-3d-chart)
