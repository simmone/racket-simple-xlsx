#lang racket

(require simple-xml)

(require rackunit/text-ui)

(require "../../../../xlsx/xlsx.rkt")
(require "../../../../writer.rkt")
(require "../../../../lib/lib.rkt")

(require rackunit "../../../../new/xl/theme/theme.rkt")

(require racket/runtime-path)
(define-runtime-path theme_file "theme.xml")

(define test-theme
  (test-suite
   "test-theme"

   (test-case
    "test-theme1"

    (parameterize 
     ([*CURRENT_XLSX* (new-xlsx)])
      (call-with-input-file theme_file
        (lambda (expected)
          (call-with-input-string
           (lists->xml (theme))
           (lambda (actual)
             (check-lines? expected actual)))))))
   ))

(run-tests test-theme)
