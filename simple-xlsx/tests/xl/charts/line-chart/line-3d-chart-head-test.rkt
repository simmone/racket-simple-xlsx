#lang racket

(require simple-xml)

(require rackunit/text-ui rackunit)

(require "../../../../lib/lib.rkt")

(require"../../../../xl/charts/line-chart.rkt")

(require racket/runtime-path)
(define-runtime-path line_3d_chart_head_file "line_3d_chart_head.xml")

(define test-line-3d-chart-head
  (test-suite
   "test-line-3d-chart-head"

   (test-case
    "test-line-3d-chart-head"

    (call-with-input-file line_3d_chart_head_file
      (lambda (expected)
        (call-with-input-string
         (lists->xml_content (line-3d-chart-head))
         (lambda (actual)
           (check-lines? expected actual))))))
   ))

(run-tests test-line-3d-chart-head)

