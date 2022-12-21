#lang racket

(require rackunit/text-ui rackunit)
(require simple-xml)

(require "../../../../lib/lib.rkt")

(require"../../../../xl/chartsheets/_rels/chartsheets-rels.rkt")

(require racket/runtime-path)
(define-runtime-path chartsheet_rels1_file "chartsheet_xml_rels1")
(define-runtime-path chartsheet_rels2_file "chartsheet_xml_rels2")

(define test-chartsheets-rels
  (test-suite
   "test-chartsheets-rels"

   (test-case
    "test-chartsheets-rels"

      (call-with-input-file chartsheet_rels1_file
        (lambda (expected)
          (call-with-input-string
           (lists->xml (chartsheets-rels 1))
           (lambda (actual)
             (check-lines? expected actual)))))

      (call-with-input-file chartsheet_rels2_file
        (lambda (expected)
          (call-with-input-string
           (lists->xml (chartsheets-rels 2))
           (lambda (actual)
             (check-lines? expected actual))))))
    ))

(run-tests test-chartsheets-rels)
