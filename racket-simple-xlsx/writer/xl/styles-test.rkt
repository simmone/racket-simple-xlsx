#lang racket

(require rackunit/text-ui)

(require rackunit "styles.rkt")

(define test-styles
  (test-suite
   "test-styles"

   (test-case
    "test-styles"
    (check-equal? (write-styles)
                  (string-append
                   "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
                   "<styleSheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"><fonts count=\"1\"><font /></fonts><fills count=\"1\"><fill /></fills><borders count=\"1\"><border /></borders><cellStyleXfs count=\"1\"><xf /></cellStyleXfs><cellXfs count=\"2\"><xf /><xf fontId=\"1\" /></cellXfs></styleSheet>")))
   ))

(run-tests test-styles)
