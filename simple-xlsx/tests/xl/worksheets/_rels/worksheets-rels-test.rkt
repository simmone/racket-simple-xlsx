#lang racket

(require simple-xml)

(require rackunit/text-ui rackunit)

(require "../../../../xlsx/xlsx.rkt")
(require "../../../../lib/lib.rkt")

(require"../../../../xl/worksheets/_rels/worksheets-rels.rkt")

(require racket/runtime-path)
(define-runtime-path worksheet_rels_file "worksheet_xml_rels")

(define test-worksheets-rels
  (test-suite
   "test-worksheets-rels"

   (test-case
    "test-worksheets-rels"

    (with-xlsx
     (lambda ()
       (add-data-sheet "Sheet1" '(("1")))
       (add-data-sheet "Sheet2" '((1)))
       (add-data-sheet "Sheet3" '((1)))
       (add-chart-sheet "Chart1" 'LINE "Chart1" '())
       (add-chart-sheet "Chart2" 'LINE "Chart2" '())
       (add-chart-sheet "Chart3" 'LINE "Chart3" '())

       (call-with-input-file worksheet_rels_file
         (lambda (expected)
           (call-with-input-string
            (lists->xml (worksheets-rels 1))
            (lambda (actual)
              (check-lines? expected actual))))))))
   ))

(run-tests test-worksheets-rels)
