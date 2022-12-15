#lang racket

(require simple-xml)

(require rackunit/text-ui rackunit)

(require "../../../../xlsx/xlsx.rkt")
(require "../../../../sheet/sheet.rkt")
(require "../../../../lib/lib.rkt")

(require"../../../../xl/charts/bar-chart.rkt")

(require racket/runtime-path)
(define-runtime-path bar_chart_head_file "bar_chart_head.xml")

(define test-bar-chart-head
  (test-suite
   "test-bar-chart-head"

   (test-case
    "test-bar-chart-head"

    (with-xlsx
     (lambda ()
       (add-data-sheet "Sheet1"
                       '(("month1" "month2" "month3" "month4" "real") (201601 100 110 1110 6.9)))
       (add-data-sheet "Sheet2" '((1)))
       (add-data-sheet "Sheet3" '((1)))
       (add-chart-sheet "Chart1" 'BAR "Chart1" '())
       (add-chart-sheet "Chart2" 'BAR "Chart2" '())
       (add-chart-sheet "Chart3" 'BAR "Chart3" '())

       (with-sheet
        (lambda ()
          (call-with-input-file bar_chart_head_file
            (lambda (expected)
              (call-with-input-string
               (lists->xml_content (bar-chart-head))
               (lambda (actual)
                 (check-lines? expected actual))))))))))
   ))

(run-tests test-bar-chart-head)

