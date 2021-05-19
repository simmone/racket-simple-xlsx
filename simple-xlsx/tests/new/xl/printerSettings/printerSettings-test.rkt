#lang racket

(require simple-xml)

(require rackunit/text-ui)

(require "../../../../xlsx/xlsx.rkt")
(require "../../../../writer.rkt")
(require "../../../../lib/lib.rkt")

(require rackunit "../../../../new/xl/printerSettings/printerSettings.rkt")

(require racket/runtime-path)
(define-runtime-path template_file "printerSettings.template")

(define test-printerSettings
  (test-suite
   "test-printerSettings"

   (test-case
    "test-printerSettings1"

    (parameterize 
     ([*CURRENT_XLSX* (new-xlsx)])

      (call-with-input-file template_file
        (lambda (expected)
          (check-equal? (port->bytes expected) (printer-settings))))))
   ))

(run-tests test-printerSettings)
