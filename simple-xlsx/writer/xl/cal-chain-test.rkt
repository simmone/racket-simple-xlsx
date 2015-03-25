#lang racket

(require rackunit/text-ui)

(require rackunit "cal-chain.rkt")

(define test-cal-chain
  (test-suite
   "test-cal-chain"

   (test-case
    "test-cal-chain"
    (check-equal? (write-cal-chain)
                  (string-append
                   "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
                   "<calcChain xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"><c r=\"C2\" i=\"1\"/></calcChain>")
    ))

   ))

(run-tests test-cal-chain)
