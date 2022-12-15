#lang racket

(require simple-xml)

(require rackunit/text-ui rackunit)

(require "../../../xlsx/xlsx.rkt")
(require "../../../lib/lib.rkt")

(require"../../../xl/printerSettings/printerSettings.rkt")

(require racket/runtime-path)
(define-runtime-path template_file "printerSettings.template")

(define test-printerSettings
  (test-suite
   "test-printerSettings"

   (test-case
    "test-printerSettings1"

    (with-xlsx
     (lambda ()
       (call-with-input-file template_file
         (lambda (expected)
           (check-equal? (port->bytes expected) (printer-settings)))))))
   ))

(run-tests test-printerSettings)
