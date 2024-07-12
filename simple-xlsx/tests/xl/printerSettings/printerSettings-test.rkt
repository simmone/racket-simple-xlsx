#lang racket

(require fast-xml
         rackunit/text-ui
         rackunit
         "../../../xlsx/xlsx.rkt"
         "../../../lib/lib.rkt"
         "../../../xl/printerSettings/printerSettings.rkt"
         racket/runtime-path)

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
