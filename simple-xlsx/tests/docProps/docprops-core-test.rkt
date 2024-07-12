#lang racket

(require fast-xml
         "../../xlsx/xlsx.rkt"
         "../../lib/lib.rkt"
         rackunit/text-ui
         rackunit
         "../../docProps/docprops-core.rkt"
         racket/runtime-path)

(define-runtime-path test_file "core_test.xml")

(define test-docprops-core
  (test-suite
   "test-docprops-core"

   (test-case
    "test-docprops-core"

    (call-with-input-file test_file
      (lambda (expected)
        (call-with-input-string
         (lists-to-xml (docprops-core (date* 44 17 13 2 1 2015 5 1 #f 28800 996159076 "CST")))
         (lambda (actual)
           (check-lines? expected actual))))))))

(run-tests test-docprops-core)
