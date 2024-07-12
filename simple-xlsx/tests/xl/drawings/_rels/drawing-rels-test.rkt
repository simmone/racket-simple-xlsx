#lang racket

(require fast-xml
         rackunit/text-ui
         rackunit
         "../../../../lib/lib.rkt"
         "../../../../xlsx/xlsx.rkt"
         "../../../../sheet/sheet.rkt"
         "../../../../xl/drawings/_rels/drawing-rels.rkt"
         racket/runtime-path)

(define-runtime-path drawing_xml_rels_file "drawing_xml_rels")

(define test-drawing-rels
  (test-suite
   "test-drawing-rels"

   (test-case
    "test-drawing-rels"

    (call-with-input-file drawing_xml_rels_file
      (lambda (expected)
        (call-with-input-string
         (lists-to-xml (drawing-rels 2))
         (lambda (actual)
           (check-lines? expected actual))))))
   ))

(run-tests test-drawing-rels)
