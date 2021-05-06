#lang racket

(require simple-xml)

(require rackunit/text-ui)

(require "../../../xlsx/xlsx.rkt")
(require "../../../writer.rkt")
(require "../../../lib/lib.rkt")

(require rackunit "../../../new/_rels/rels.rkt")

(require racket/runtime-path)
(define-runtime-path rels_file "rels")

(define test-rels
  (test-suite
   "test-rels"

   (test-case
    "test-rels"

    (parameterize 
     ([*CURRENT_XLSX* (new-xlsx)])
      (call-with-input-file rels_file
        (lambda (expected)
          (call-with-input-string
           (lists->xml (rels))
           (lambda (actual)
             (check-lines? expected actual)))))))
   ))

(run-tests test-rels)
