#lang racket

(require rackunit/text-ui
         rackunit
         "../../style/number-style.rkt")

(define test-number-style
  (test-suite
   "test-number-style"

   (test-case
    "test-update-number-style"

    (let ([number_style (NUMBER-STYLE "164" "0.00")]
          [new_style (NUMBER-STYLE "165" "%0.00")])
      (update-number-style number_style new_style)
      (check-equal? number_style (NUMBER-STYLE "164" "%0.00")))
    )

   ))

(run-tests test-number-style)
