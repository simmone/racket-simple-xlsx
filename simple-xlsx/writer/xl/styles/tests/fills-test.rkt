#lang racket

(require rackunit/text-ui)

(require "../../../../lib/lib.rkt")

(require rackunit "../styles.rkt")

(require racket/runtime-path)
(define-runtime-path test_file "fills-test.dat")

(define test-styles
  (test-suite
   "test-styles"

   (test-case
    "test-color-style"
    
    (let ([fill_list (list #hash((fgColor . "FF0000")) #hash((fgColor . "00FF00")) #hash((fgColor . "0000FF")))])
      
      (call-with-input-file test_file
        (lambda (expected)
          (call-with-input-string
           (write-fills fill_list)
           (lambda (actual)
             (check-lines? expected actual)))))
      ))))

(run-tests test-styles)
