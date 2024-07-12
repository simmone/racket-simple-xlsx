#lang racket

(require fast-xml
         rackunit/text-ui
         rackunit
         "../../../../xlsx/xlsx.rkt"
         "../../../../sheet/sheet.rkt"
         "../../../../lib/lib.rkt"
         "../../../../xl/charts/pie-chart.rkt"
         racket/runtime-path)

(define-runtime-path pie_chart_head_file "pie_chart_head.xml")

(define test-pie-chart-head
  (test-suite
   "test-pie-chart-head"

   (test-case
    "test-pie-chart-head"

    (with-xlsx
     (lambda ()
       (add-data-sheet "Sheet1"
                       '(("month1" "month2" "month3" "month4" "real") (201601 100 110 1110 6.9)))
       (add-data-sheet "Sheet2" '((1)))
       (add-data-sheet "Sheet3" '((1)))
       (add-chart-sheet "Chart1" 'PIE "Chart1" '())
       (add-chart-sheet "Chart2" 'PIE "Chart2" '())
       (add-chart-sheet "Chart3" 'PIE "Chart3" '())

       (call-with-input-file pie_chart_head_file
         (lambda (expected)
           (call-with-input-string
            (lists-to-xml_content (pie-chart-head))
            (lambda (actual)
              (check-lines? expected actual))))))))
   ))

(run-tests test-pie-chart-head)

