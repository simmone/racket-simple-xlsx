#lang racket

(require simple-xml)

(require rackunit/text-ui rackunit)

(require "../../../../lib/lib.rkt")

(require"../../../../xl/drawings/drawing.rkt")

(require racket/runtime-path)
(define-runtime-path drawing_file "drawing.xml")

(define test-drawing
  (test-suite
   "test-drawing"

   (test-case
    "test-drawing"

    (call-with-input-file drawing_file
      (lambda (expected)
        (call-with-input-string
         (lists->xml (drawing 2))
         (lambda (actual)
           (check-lines? expected actual))))))
   ))

(run-tests test-drawing)
