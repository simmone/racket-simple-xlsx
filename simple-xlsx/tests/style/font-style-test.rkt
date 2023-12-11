#lang racket

(require rackunit/text-ui rackunit)

(require "../../style/font-style.rkt")

(define test-font-style
  (test-suite
   "test-font-style"

   (test-case
    "test-update-font-style"

    (let ([font_style (FONT-STYLE 10 "Arial" "0000FF")]
          [new_style (FONT-STYLE 11.0 "Italic" "0000F0")]
          )

      (update-font-style font_style new_style)
      (check-equal? font_style (FONT-STYLE 11.0 "Italic" "0000F0"))
      )
    )

   ))

(run-tests test-font-style)
