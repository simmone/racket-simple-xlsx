#lang racket

(require simple-xml)

(require rackunit/text-ui rackunit)

(require "../../../../xlsx/xlsx.rkt")
(require "../../../../sheet/sheet.rkt")
(require "../../../../lib/lib.rkt")

(require"../../../../xl/charts/charts.rkt")

(require racket/runtime-path)
(define-runtime-path plotArea_head_file "plotArea_head.xml")

(define test-plotArea
  (test-suite
   "test-plotArea"

   (test-case
    "test-plotArea-head"

    (with-xlsx
     (lambda ()
       (add-data-sheet "Sheet1"
                       '(("month1" "month2" "month3" "month4" "real") (201601 100 110 1110 6.9)))
       (add-data-sheet "Sheet2" '((1)))
       (add-data-sheet "Sheet3" '((1)))

       (call-with-input-file plotArea_head_file
         (lambda (expected)
           (call-with-input-string
            (lists->xml_content (to-plotArea-head))
            (lambda (actual)
              (check-lines? expected actual))))))))
   ))

(run-tests test-plotArea)
