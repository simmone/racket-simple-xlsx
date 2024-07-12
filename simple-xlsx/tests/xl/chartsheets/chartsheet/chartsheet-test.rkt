#lang racket

(require fast-xml
         rackunit/text-ui
         rackunit
         "../../../../lib/lib.rkt"
         "../../../../xl/chartsheets/chartsheet.rkt"
         racket/runtime-path)

(define-runtime-path chartsheet_file "chart_sheet.xml")

(define test-chart-sheet
  (test-suite
   "test-chart-sheet"

   (test-case
    "test-chart-sheet"

    (call-with-input-file chartsheet_file
      (lambda (expected)
        (call-with-input-string
         (lists-to-xml (chart-sheet 1))
         (lambda (actual)
           (check-lines? expected actual))))))
   ))


(run-tests test-chart-sheet)
