#lang racket

(require simple-xml)

(require rackunit/text-ui)

(require "../../../../lib/lib.rkt")

(require rackunit "../../../../new/xl/styles/styles.rkt")

(require racket/runtime-path)
(define-runtime-path test_file "empty-test.xml")

(define test-styles
  (test-suite
   "test-styles"

   (test-case
    "test-empty-style"
    
    (call-with-input-file test_file
      (lambda (expected)
        (call-with-input-string
         (lists->xml (styles '() '() '() '() '()))
         (lambda (actual)
           (check-lines? expected actual))))))
   ))

(run-tests test-styles)
