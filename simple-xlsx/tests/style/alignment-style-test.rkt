#lang racket

(require rackunit/text-ui rackunit)

(require "../../style/alignment-style.rkt")

(define test-alignment-style
  (test-suite
   "test-alignment-style"

   (test-case
    "test-update-alignment-style"

    (let ([alignment_style (ALIGNMENT-STYLE "right" "top")]
          [new1_style (ALIGNMENT-STYLE "left" "")]
          [new2_style (ALIGNMENT-STYLE "" "bottom")]
          )
      (update-alignment-style alignment_style new1_style)
      (check-equal? alignment_style (ALIGNMENT-STYLE "left" "top"))

      (update-alignment-style alignment_style new2_style)
      (check-equal? alignment_style (ALIGNMENT-STYLE "left" "bottom"))
      ))
   ))

(run-tests test-alignment-style)
