#lang racket

(require rackunit/text-ui rackunit)

(require "../../../style/set-styles.rkt")
(require "../../../style/border-style.rkt")
(require "../../../style/font-style.rkt")
(require "../../../style/alignment-style.rkt")
(require "../../../style/number-style.rkt")
(require "../../../style/fill-style.rkt")
(require "../../../style/style.rkt")

(struct T_STYLE (
                 (a #:mutable)
                 (b #:mutable)
                 )
        #:transparent
        )

(struct P_STYLE (
                 (c #:mutable)
                 (d #:mutable)
                 )
        #:transparent
        )

(define test-styles
  (test-suite
   "test-styles"

   (test-case
    "test-update-hash-struct"

    (let ([style (T_STYLE "1" "2")]
          [data_hash (make-hash)])

      (hash-set! data_hash 1 style)
      (check-equal? style (hash-ref data_hash 1))

      (set-T_STYLE-a! (hash-ref data_hash 1) "3")
      (check-equal? (T_STYLE "3" "2") (hash-ref data_hash 1))

      (let ([t_style (hash-ref data_hash 1)])
        (set-T_STYLE-b! (hash-ref data_hash 1) "4"))
      (check-equal? (T_STYLE "3" "4") (hash-ref data_hash 1))

      ))

   (test-case
    "test-update-recursive-struct"

    (let ([style (P_STYLE
                  "1"
                  (T_STYLE "2" "3"))])

      (let ([t_style (P_STYLE-d style)])
        (set-T_STYLE-b! t_style "4"))


      (check-equal? (P_STYLE "1" (T_STYLE "2" "4")) style)
      ))

   ))

(run-tests test-styles)
