#lang racket

(require simple-xml)

(require rackunit/text-ui rackunit)

(require "../../../../xlsx/xlsx.rkt")
(require "../../../../sheet/sheet.rkt")
(require "../../../../lib/lib.rkt")
(require"../../../../xl/charts/bar-chart.rkt")
(require"../../../../xl/charts/lib.rkt")

(require racket/runtime-path)
(define-runtime-path bar_chart_file "bar_chart.xml")

(define test-bar-chart
  (test-suite
   "test-bar-chart"

   (test-case
    "test-bar-chart"

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
       (add-chart-sheet "Chart1" 'BAR "Chart1" '())
       (add-chart-sheet "Chart2" 'BAR "Chart2" '())
       (add-chart-sheet "Chart3" 'BAR "Chart3" '())

       (with-sheet
        (lambda ()
          (call-with-input-file bar_chart_file
            (lambda (expected)
              (call-with-input-string
               (lists->xml_content
                (to-bar-chart-sers
                 '(
                   ("CAT" "DataSheet" "B1-D1" "DataSheet" "B2-D2")
                   ("Puma" "DataSheet" "B1-D1" "DataSheet" "B3-D3")
                   ("Brooks" "DataSheet" "B1-D1" "DataSheet" "B4-D4")
                  )))
               (lambda (actual)
                 (check-lines? expected actual))))))))))
   ))

(run-tests test-bar-chart)
