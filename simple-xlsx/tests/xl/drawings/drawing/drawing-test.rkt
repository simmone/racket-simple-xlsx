#lang racket

(require fast-xml
         rackunit/text-ui
         rackunit
         "../../../../lib/lib.rkt"
         "../../../../xl/drawings/drawing.rkt"
         racket/runtime-path)

(define-runtime-path drawing_file "drawing.xml")

(define test-drawing
  (test-suite
   "test-drawing"

   (test-case
    "test-drawing"

    (call-with-input-file drawing_file
      (lambda (expected)
        (call-with-input-string
         (lists-to-xml (drawing 2))
         (lambda (actual)
           (check-lines? expected actual))))))
   ))

(run-tests test-drawing)
