#lang racket

(require simple-xml)

(require "../../xlsx/xlsx.rkt")
(require "../../lib/lib.rkt")

(require rackunit/text-ui rackunit)

(require"../../docProps/docprops-core.rkt")

(require racket/runtime-path)
(define-runtime-path test_file "core_test.xml")

(define test-docprops-core
  (test-suite
   "test-docprops-core"

   (test-case
    "test-docprops-core"

    (call-with-input-file test_file
      (lambda (expected)
        (call-with-input-string
         (lists->xml (docprops-core (date* 44 17 13 2 1 2015 5 1 #f 28800 996159076 "CST")))
         (lambda (actual)
           (check-lines? expected actual))))))))

(run-tests test-docprops-core)
