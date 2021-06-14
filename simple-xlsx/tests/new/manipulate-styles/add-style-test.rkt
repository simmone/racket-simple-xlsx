#lang racket

(require rackunit/text-ui)

(require "../../../lib/lib.rkt")
(require "../../../sheet/sheet.rkt")

(require rackunit "../../../xlsx/xlsx.rkt")

(define test-add-style
  (test-suite
   "test-add-style"

   (test-case
    "test-add-style"

    (parameterize 
     ([*CURRENT_XLSX* (new-xlsx)])

     (add-data-sheet "Sheet1" '(("chenxiao" "love" "陈思衡")))
     
     (with-sheet
      "Sheet1"
      (lambda ()
        (add-cell-style "A1" '( (fontSize . 20) (fontName . "Impact")))
        (let ([style_hash (XLSX-style_hash (*CURRENT_XLSX*))]
              [font_hash (XLSX-font_hash (*CURRENT_XLSX*))])
        ))

     ))))
    
(run-tests test-styles)
