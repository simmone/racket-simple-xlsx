#lang racket

(require fast-xml
         rackunit/text-ui
         rackunit
         "../../../../xlsx/xlsx.rkt"
         "../../../../sheet/sheet.rkt"
         "../../../../lib/lib.rkt"
         "../../../../xl/charts/line-chart.rkt"
         racket/runtime-path)

(define-runtime-path line_chart_head_file "line_chart_head.xml")

(define test-line-chart-head
  (test-suite
   "test-line-chart-head"

   (test-case
    "test-line-chart-head"

    (with-xlsx
     (lambda ()
       (add-data-sheet "Sheet1"
                       '(("month1" "month2" "month3" "month4" "real") (201601 100 110 1110 6.9)))
       (add-data-sheet "Sheet2" '((1)))
       (add-data-sheet "Sheet3" '((1)))
       (add-chart-sheet "Chart1" 'LINE "Chart1" '())
       (add-chart-sheet "Chart2" 'LINE "Chart2" '())
       (add-chart-sheet "Chart3" 'LINE "Chart3" '())

       (call-with-input-file line_chart_head_file
         (lambda (expected)
           (call-with-input-string
            (lists-to-xml_content (line-chart-head))
            (lambda (actual)
              (check-lines? expected actual))))))))
  ))

(run-tests test-line-chart-head)

