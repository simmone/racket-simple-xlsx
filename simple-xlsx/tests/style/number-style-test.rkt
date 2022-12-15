#lang racket

(require rackunit/text-ui rackunit)

(require "../../style/number-style.rkt")

(define test-number-style
  (test-suite
   "test-number-style"

   (test-case
    "test-set-number-style1"

    (let ([number_style (number-style-from-hash-code "0.00")])
      (check-equal? (NUMBER-STYLE-format number_style) "0.00")
      (check-equal? (NUMBER-STYLE-hash_code number_style) "0.00"))

    (let ([number_style (number-style-from-hash-code "")])
      (check-false number_style)))

   (test-case
    "test-number-style<=?"

    (let ([number1_style #f]
          [number2_style (number-style-from-hash-code "0.01")])
      (check-true (number-style<? number1_style number2_style))
      (check-false (number-style=? number1_style number2_style)))

    (let ([number1_style (number-style-from-hash-code "0.00")]
          [number2_style #f])
      (check-false (number-style<? number1_style number2_style))
      (check-false (number-style=? number1_style number2_style)))

    (let ([number1_style #f]
          [number2_style #f])
      (check-false (number-style<? number1_style number2_style))
      (check-true (number-style=? number1_style number2_style)))

    (let ([number1_style (number-style-from-hash-code "0.00")]
          [number2_style (number-style-from-hash-code "0.00")])
      (check-false (number-style<? number1_style number2_style))
      (check-true (number-style=? number1_style number2_style)))

    (let ([number1_style (number-style-from-hash-code "0.00")]
          [number2_style (number-style-from-hash-code "0.01")])
      (check-true (number-style<? number1_style number2_style))
      (check-false (number-style=? number1_style number2_style)))
    )
   ))

(run-tests test-number-style)
