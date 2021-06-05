#lang racket

(require simple-xml)

(require rackunit/text-ui)

(require "../../../../lib/lib.rkt")

(require rackunit "../../../../new/xl/styles/styles.rkt")

(require racket/runtime-path)
(define-runtime-path test_file "fonts-test.xml")

(define test-styles
  (test-suite
   "test-styles"

   (test-case
    "test-color-style"
    
    (let ([font_list 
           (list 
            #hash((fontSize . 20)) 
            #hash((fontSize . 30) (fontColor . "FF0000") (fontName . "Impact"))
            #hash((fontSize . 40) (fontColor . "0000FF"))
            )])
      
      (call-with-input-file test_file
        (lambda (expected)
          (call-with-input-string
           (lists->xml_content (fonts font_list))
           (lambda (actual)
             (check-lines? expected actual)))))
      ))))

(run-tests test-styles)
