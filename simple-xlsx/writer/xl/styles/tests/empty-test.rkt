#lang racket

(require rackunit/text-ui)

(require "../../../../lib/lib.rkt")

(require rackunit "../styles.rkt")

(require racket/runtime-path)
(define-runtime-path test_file "empty-test.dat")

(define test-styles
  (test-suite
   "test-styles"

   (test-case
    "test-empty-style"
    
    (call-with-input-file test_file
      (lambda (expected)
        (call-with-input-string
         (write-styles '() '() '() '() '())
         (lambda (actual)
           (check-lines? expected actual))))))
   ))

(run-tests test-styles)
