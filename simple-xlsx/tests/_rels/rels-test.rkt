#lang racket

(require fast-xml
         rackunit/text-ui
         rackunit
         "../../xlsx/xlsx.rkt"
         "../../lib/lib.rkt"
         "../../_rels/rels.rkt"
         racket/runtime-path)

(define-runtime-path rels_file "rels")

(define test-rels
  (test-suite
   "test-rels"

   (test-case
    "test-rels"

    (call-with-input-file rels_file
      (lambda (expected)
        (call-with-input-string
         (lists-to-xml (rels))
         (lambda (actual)
           (check-lines? expected actual))))))
   ))

(run-tests test-rels)
