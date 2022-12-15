#lang racket

(require rackunit/text-ui rackunit)

(require"../../style/lib.rkt")

(define test-lib
  (test-suite
   "test-lib"

   (test-case
    "test-squeeze-range-hash"

    (let ([range_hash (make-hash)])
      (hash-set! range_hash 1 5)
      (hash-set! range_hash 2 5)
      (hash-set! range_hash 3 6)
      (hash-set! range_hash 5 6)

      (let ([squeeze_list (squeeze-range-hash range_hash)])
        (check-equal? (length squeeze_list) 3)
        (check-equal? (list-ref squeeze_list 0) '(1 2 5))
        (check-equal? (list-ref squeeze_list 1) '(3 3 6))
        (check-equal? (list-ref squeeze_list 2) '(5 5 6))))

    (let ([range_hash (make-hash)])
      (let ([squeeze_list (squeeze-range-hash range_hash)])
        (check-equal? (length squeeze_list) 0)))

    (let ([range_hash (make-hash)])
      (hash-set! range_hash 2 4)
      (let ([squeeze_list (squeeze-range-hash range_hash)])
        (check-equal? (length squeeze_list) 1)
        (check-equal? (list-ref squeeze_list 0) '(2 2 4))
        ))

    (let ([range_hash (make-hash)])
      (hash-set! range_hash 1 5)
      (hash-set! range_hash 2 5)
      (hash-set! range_hash 3 6)
      (hash-set! range_hash 4 7)
      (hash-set! range_hash 5 6)
      (hash-set! range_hash 6 6)

      (let ([squeeze_list (squeeze-range-hash range_hash)])
        (check-equal? (length squeeze_list) 4)
        (check-equal? (list-ref squeeze_list 0) '(1 2 5))
        (check-equal? (list-ref squeeze_list 1) '(3 3 6))
        (check-equal? (list-ref squeeze_list 2) '(4 4 7))
        (check-equal? (list-ref squeeze_list 3) '(5 6 6))))
    )

   (test-case
    "test-rgb?"

    (check-true (rgb? "123456"))
    (check-false (rgb? "1234567"))
    (check-false (rgb? "12345"))
    (check-false (rgb? "1234(6"))
    (check-true (rgb? "11aaBB"))
    (check-true (rgb? "11aaBBAA"))
    (check-true (rgb? "11aaBB00"))
    )

   ))

(run-tests test-lib)
