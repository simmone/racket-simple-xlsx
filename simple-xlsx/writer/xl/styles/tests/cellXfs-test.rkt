#lang racket

(require rackunit/text-ui)

(require racket/runtime-path)
(define-runtime-path test_file "cellXfs-test.dat")

(require "../../../../lib/lib.rkt")

(require rackunit "../styles.rkt")

(define test-styles
  (test-suite
   "test-styles"

   (test-case
    "test-cellXfs"
    
    (let ([style_list 
           (list 
            #hash((fill . 1)) 
            #hash((fill . 2) (font . 1)) 
            #hash((font . 2))
            #hash((border . 1))
            )])
      
      (call-with-input-file test_file
        (lambda (expected)
          (call-with-input-string
           (write-cellXfs style_list)
           (lambda (actual)
             (check-lines? expected actual)))))
      ))))

(run-tests test-styles)
