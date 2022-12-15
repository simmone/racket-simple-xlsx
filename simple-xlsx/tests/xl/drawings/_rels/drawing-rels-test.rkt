#lang racket

(require simple-xml)

(require rackunit/text-ui rackunit)

(require "../../../../lib/lib.rkt")
(require "../../../../xlsx/xlsx.rkt")
(require "../../../../sheet/sheet.rkt")

(require"../../../../xl/drawings/_rels/drawing-rels.rkt")

(require racket/runtime-path)
(define-runtime-path drawing_xml_rels_file "drawing_xml_rels")

(define test-drawing-rels
  (test-suite
   "test-drawing-rels"

   (test-case
    "test-drawing-rels"

    (call-with-input-file drawing_xml_rels_file
      (lambda (expected)
        (call-with-input-string
         (lists->xml (drawing-rels 2))
         (lambda (actual)
           (check-lines? expected actual))))))
   ))

(run-tests test-drawing-rels)
