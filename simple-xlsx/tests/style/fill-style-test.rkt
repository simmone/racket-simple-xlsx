#lang racket

(require rackunit/text-ui rackunit)

(require"../../style/fill-style.rkt")

(define test-fill-style
  (test-suite
   "test-fill-style"

   (test-case
    "test-set-fill-style1"

    (let ([fill_style (fill-style-from-hash-code "0000ff<p>solid")])
      (check-equal? (FILL-STYLE-color fill_style) "0000FF")
      (check-equal? (FILL-STYLE-pattern fill_style) "solid")
      (check-equal? (FILL-STYLE-hash_code fill_style) "0000FF<p>solid"))

    (let ([fill_style (fill-style-from-hash-code "")])
      (check-false fill_style)))

   (test-case
    "test-fill-style<?"

    (let ([fill1_style #f]
          [fill2_style (fill-style-from-hash-code "0000FF<p>solid")])
      (check-true (fill-style<? fill1_style fill2_style))
      (check-false (fill-style=? fill1_style fill2_style)))

    (let ([fill1_style (fill-style-from-hash-code "0000F0<p>solid")]
          [fill2_style #f])
      (check-false (fill-style<? fill1_style fill2_style))
      (check-false (fill-style=? fill1_style fill2_style)))

    (let ([fill1_style #f]
          [fill2_style #f])
      (check-false (fill-style<? fill1_style fill2_style))
      (check-true (fill-style=? fill1_style fill2_style)))

    (let ([fill1_style (fill-style-from-hash-code "0000FF<p>solid")]
          [fill2_style (fill-style-from-hash-code "0000FF<p>solid")])
      (check-false (fill-style<? fill1_style fill2_style))
      (check-true (fill-style=? fill1_style fill2_style)))

    (let ([fill1_style (fill-style-from-hash-code "0000F0<p>solid")]
          [fill2_style (fill-style-from-hash-code "0000FF<p>solid")])
      (check-true (fill-style<? fill1_style fill2_style))
      (check-false (fill-style=? fill1_style fill2_style)))

    (let ([fill1_style (fill-style-from-hash-code "0000FF<p>gray125")]
          [fill2_style (fill-style-from-hash-code "0000FF<p>solid")])
      (check-true (fill-style<? fill1_style fill2_style)))
    )

   ))

(run-tests test-fill-style)
