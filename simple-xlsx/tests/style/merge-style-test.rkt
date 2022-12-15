#lang racket

(require rackunit/text-ui rackunit)

(require "../../style/merge-style.rkt")

(define test-merge-style
  (test-suite
   "test-merge-style"

   (test-case
    "test-set-merge-style1"

    (let ([merge_style (merge-style-from-hash-code "A1:C3")])
      (check-equal? (MERGE-STYLE-cell_range merge_style) "A1:C3")
      (check-equal? (MERGE-STYLE-hash_code merge_style) "A1:C3"))

    (let ([merge_style (merge-style-from-hash-code "")])
      (check-false merge_style)))

   (test-case
    "test-merge-style<=?"

    (let ([merge1_style #f]
          [merge2_style (merge-style-from-hash-code "A1:C4")])
      (check-true (merge-style<? merge1_style merge2_style))
      (check-false (merge-style=? merge1_style merge2_style)))

    (let ([merge1_style (merge-style-from-hash-code "A1:C3")]
          [merge2_style #f])
      (check-false (merge-style<? merge1_style merge2_style))
      (check-false (merge-style=? merge1_style merge2_style)))

    (let ([merge1_style #f]
          [merge2_style #f])
      (check-false (merge-style<? merge1_style merge2_style))
      (check-true (merge-style=? merge1_style merge2_style)))

    (let ([merge1_style (merge-style-from-hash-code "A1:C3")]
          [merge2_style (merge-style-from-hash-code "A1:C3")])
      (check-false (merge-style<? merge1_style merge2_style))
      (check-true (merge-style=? merge1_style merge2_style)))

    (let ([merge1_style (merge-style-from-hash-code "A1:C3")]
          [merge2_style (merge-style-from-hash-code "A1:C4")])
      (check-true (merge-style<? merge1_style merge2_style))
      (check-false (merge-style=? merge1_style merge2_style)))
    )
   ))

(run-tests test-merge-style)
