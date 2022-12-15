#lang racket

(require simple-xml)

(require rackunit/text-ui rackunit)

(require "../../../../lib/lib.rkt")

(require"../../../../xl/chartsheets/chartsheet.rkt")

(require racket/runtime-path)
(define-runtime-path chartsheet_file "chart_sheet.xml")

(define test-chart-sheet
  (test-suite
   "test-chart-sheet"

   (test-case
    "test-chart-sheet"

    (call-with-input-file chartsheet_file
      (lambda (expected)
        (call-with-input-string
         (lists->xml (chart-sheet 1))
         (lambda (actual)
           (check-lines? expected actual))))))
   ))


(run-tests test-chart-sheet)
