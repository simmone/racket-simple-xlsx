#lang racket

(require fast-xml
         rackunit/text-ui
         rackunit
         "../../../xlsx/xlsx.rkt"
         "../../../lib/lib.rkt"
         "../../../xl/printerSettings/printerSettings.rkt"
         racket/runtime-path)

(define-runtime-path printerSettings_file1 "printerSettings1.bin")
(define-runtime-path printerSettings_file2 "printerSettings2.bin")
(define-runtime-path printerSettings_file3 "printerSettings3.bin")

(define test-printerSettings
  (test-suite
   "test-printerSettings"

   (test-case
    "test-write-printer-settings"

    (with-xlsx
     (lambda ()
       (add-data-sheet "Sheet1" '((1)))
       (add-data-sheet "Sheet2" '((1)))
       (add-data-sheet "Sheet3" '((1)))

       (dynamic-wind
           (lambda ()
             (write-printer-settings (apply build-path (drop-right (explode-path printerSettings_file1) 1))))
           (lambda ()
             (call-with-input-file printerSettings_file1
               (lambda (actual1)
                 (check-equal? (printer-settings) (port->bytes actual1))))

             (call-with-input-file printerSettings_file2
               (lambda (actual2)
                 (check-equal? (printer-settings) (port->bytes actual2))))

             (call-with-input-file printerSettings_file3
               (lambda (actual3)
                 (check-equal? (printer-settings) (port->bytes actual3)))))
           (lambda ()
             (when (file-exists? printerSettings_file1) (delete-file printerSettings_file1))
             (when (file-exists? printerSettings_file2) (delete-file printerSettings_file2))
             (when (file-exists? printerSettings_file3) (delete-file printerSettings_file3)))))))
   ))

(run-tests test-printerSettings)
