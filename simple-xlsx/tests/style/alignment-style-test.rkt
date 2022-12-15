#lang racket

(require rackunit/text-ui rackunit)

(require "../../style/alignment-style.rkt")

(define test-alignment-style
  (test-suite
   "test-alignment-style"

   (test-case
    "test-set-alignment-style"

    (let ([alignment_style (alignment-style-from-hash-code "left<p>top")])
      (check-equal? (ALIGNMENT-STYLE-horizontal_placement alignment_style) "left")
      (check-equal? (ALIGNMENT-STYLE-vertical_placement alignment_style) "top")
      (check-equal? (ALIGNMENT-STYLE-hash_code alignment_style) "left<p>top"))

    (let ([alignment_style (alignment-style-from-hash-code "")])
      (check-false alignment_style))
    )

   (test-case
    "test-alignment-style<=?"

    (let ([alignment1_style #f]
          [alignment2_style (alignment-style-from-hash-code "right<p>top")])
      (check-true (alignment-style<? alignment1_style alignment2_style))
      (check-false (alignment-style=? alignment1_style alignment2_style)))

    (let ([alignment1_style (alignment-style-from-hash-code "left<p>top")]
          [alignment2_style #f])
      (check-false (alignment-style<? alignment1_style alignment2_style))
      (check-false (alignment-style=? alignment1_style alignment2_style)))

    (let ([alignment1_style #f]
          [alignment2_style #f])
      (check-false (alignment-style<? alignment1_style alignment2_style))
      (check-true (alignment-style=? alignment1_style alignment2_style)))

    (let ([alignment1_style (alignment-style-from-hash-code "right<p>top")]
          [alignment2_style (alignment-style-from-hash-code "right<p>top")])
      (check-false (alignment-style<? alignment1_style alignment2_style))
      (check-true (alignment-style=? alignment1_style alignment2_style)))

    (let ([alignment1_style (alignment-style-from-hash-code "left<p>top")]
          [alignment2_style (alignment-style-from-hash-code "right<p>top")])
      (check-true (alignment-style<? alignment1_style alignment2_style))
      (check-false (alignment-style=? alignment1_style alignment2_style)))

    (let ([alignment1_style (alignment-style-from-hash-code "right<p>bottom")]
          [alignment2_style (alignment-style-from-hash-code "right<p>top")])
      (check-true (alignment-style<? alignment1_style alignment2_style)))
    )
   ))

(run-tests test-alignment-style)
