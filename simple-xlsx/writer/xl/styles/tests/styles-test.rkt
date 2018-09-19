#lang racket

(require rackunit/text-ui)

(require "../../../../lib/lib.rkt")

(require rackunit "../styles.rkt")

(require racket/runtime-path)
(define-runtime-path test_file "styles-test.dat")

(define test-styles
  (test-suite
   "test-styles"

   (test-case
    "test-styles"
    
    (let ([style_list 
           (list 
            #hash((fill . 1) (font . 3) (numFmt . 165))
            #hash((fill . 2))
            #hash((fill . 3) (font . 1) (numFmt . 166))
            #hash((font . 2))
            )
           ]
          [fill_list 
           (list 
            #hash((fgColor . "FF0000")) 
            #hash((fgColor . "00FF00")) 
            #hash((fgColor . "0000FF")))]
          [font_list 
           (list 
            #hash((fontSize . 20)) 
            #hash((fontSize . 30) (fontColor . "red"))
            #hash((fontSize . 40))
            )]
          [numFmt_list
           (list
            #hash((numberPrecision . 2))
            #hash((numberPrecision . 2) (numberPercent . #t))
            )]
          [border_list
           (list
            #hash((borderStyle . thick))
            #hash((borderColor . "red"))
            )]
          )

      (call-with-input-file test_file
        (lambda (expected)
          (call-with-input-string
           (write-styles style_list fill_list font_list numFmt_list border_list)
           (lambda (actual)
             (check-lines? expected actual)))))
      ))))

(run-tests test-styles)
