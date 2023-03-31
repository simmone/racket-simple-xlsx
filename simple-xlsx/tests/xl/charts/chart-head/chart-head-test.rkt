#lang racket

(require simple-xml)

(require rackunit/text-ui rackunit)

(require "../../../../xlsx/xlsx.rkt")
(require "../../../../sheet/sheet.rkt")
(require "../../../../lib/lib.rkt")

(require"../../../../xl/charts/charts.rkt")

(require racket/runtime-path)
(define-runtime-path chart_head_file "chart_head.xml")

(define test-chart
  (test-suite
   "test-chart"

   (test-case
    "test-chart-head"

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
          (call-with-input-file chart_head_file
            (lambda (expected)
              (call-with-input-string
               (lists->xml_content (to-chart-head))
               (lambda (actual)
                 (check-lines? expected actual))))))))))
   ))

(run-tests test-chart)