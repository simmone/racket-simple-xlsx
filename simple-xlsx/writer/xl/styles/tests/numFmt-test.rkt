#lang racket

(require rackunit/text-ui)

(require "../../../../lib/lib.rkt")

(require rackunit "../styles.rkt")

(require racket/runtime-path)
(define-runtime-path test_file "numFmt-test.dat")

(define test-styles
  (test-suite
   "test-styles"

   (test-case
    "test-get-numFormatCode"

    (let ([format_hash '#hash( (numberPrecision . 2) )])
      (check-equal? (get-numFormatCode format_hash) "0.00"))

    (let ([format_hash (make-hash)])
      (check-equal? (get-numFormatCode format_hash) "0.00"))

    (let ([format_hash '#hash( (numberPrecision . 3) )])
      (check-equal? (get-numFormatCode format_hash) "0.000"))

    (let ([format_hash '#hash( (numberPrecision . 4) (numberThousands . #t) )])
      (check-equal? (get-numFormatCode format_hash) "#,###0.0000"))

    (let ([format_hash '#hash( (numberPrecision . 3) (numberPercent . #t) )])
      (check-equal? (get-numFormatCode format_hash) "0.000%"))

    (let ([format_hash '#hash( (numberPrecision . 0) )])
      (check-equal? (get-numFormatCode format_hash) "0"))

    (let ([format_hash '#hash( (numberPrecision . -1) )])
      (check-equal? (get-numFormatCode format_hash) "0.00"))

    (let ([format_hash '#hash( (numberPrecision . 1.232132) )])
      (check-equal? (get-numFormatCode format_hash) "0.00"))
    )

   (test-case
    "test-numFmt"

    (let (
          [numFmt_list
           (list
            #hash((numberPrecision . 2))
            #hash((numberPrecision . 2) (numberThousands . #t))
            #hash((numberPrecision . 2) (numberPercent . #t))
            #hash((numberPrecision . 0))
            #hash((numberPrecision . 3))
            #hash((dateFormat . "yyyy年mm月dd日"))
            #hash((dateFormat . "yyyy-mm-dd"))
            #hash((dateFormat . "yyyy/mm/dd"))
            )]
          )

      (call-with-input-file test_file
        (lambda (expected)
          (call-with-input-string
           (write-numFmts numFmt_list)
           (lambda (actual)
             (check-lines? expected actual)))))
      ))))

(run-tests test-styles)
