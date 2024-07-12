#lang racket

(require rackunit/text-ui
         rackunit
         "../../style/fill-style.rkt")

(define test-fill-style
  (test-suite
   "test-fill-style"

   (test-case
    "test-update-fill-style"

    (let ([fill_style (FILL-STYLE "0000FF" "solid")]
          [new1_style (FILL-STYLE "000FFF" "gray125")])

      (update-fill-style fill_style new1_style)

      (check-equal? fill_style (FILL-STYLE "000FFF" "gray125")))
    )

   ))

(run-tests test-fill-style)
