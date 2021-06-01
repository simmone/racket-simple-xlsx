#lang racket

(require simple-xml)

(require rackunit/text-ui)

(require "../../../../lib/lib.rkt")

(require rackunit "../../../../new/xl/styles/styles.rkt")

(require racket/runtime-path)
(define-runtime-path test_file "borders-test.xml")

(define test-styles
  (test-suite
   "test-styles"

   (test-case
    "test-border-style"
    
    (let ([border_list 
           (list 
            '#hash((borderDirection . all) (borderStyle . double) (borderColor . "000000"))
            '#hash((borderDirection . all) (borderStyle . dashed) (borderColor . "0000FF"))
            '#hash((borderDirection . left) (borderStyle . thick) (borderColor . "0000FF"))
            '#hash((borderDirection . right) (borderStyle . thick) (borderColor . "0000FF"))
            '#hash((borderDirection . top) (borderStyle . thick) (borderColor . "0000FF"))
            '#hash((borderDirection . bottom) (borderStyle . thick) (borderColor . "0000FF"))
            )])
      
      (printf "~a\n" (lists->xml_content (borders border_list)))
      
      (call-with-input-file test_file
        (lambda (expected)
          (call-with-input-string
           (lists->xml_content (borders border_list))
           (lambda (actual)
             (check-lines? expected actual)))))
      ))))

(run-tests test-styles)
