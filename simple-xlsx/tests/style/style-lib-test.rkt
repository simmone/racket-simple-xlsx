#lang racket

(require rackunit/text-ui
         rackunit
         "../../style/style-lib.rkt")

(define test-lib
  (test-suite
   "test-lib"

   (test-case
    "test-rgb?"

    (check-true (rgb? "123456"))
    (check-false (rgb? "1234567"))
    (check-false (rgb? "12345"))
    (check-false (rgb? "1234(6"))
    (check-false (rgb? "11aaBB"))
    (check-true (rgb? "00FF00"))
    (check-true (rgb? "11AABBAA"))
    (check-true (rgb? "11AABB00"))
    )

   ))

(run-tests test-lib)
