#lang racket

(require fast-xml
         rackunit/text-ui
         rackunit
         "../../../xlsx/xlsx.rkt"
         "../../../lib/lib.rkt"
         "../../../xl/theme/theme.rkt"
         racket/runtime-path)

(define-runtime-path theme_file "theme.xml")

(define test-theme
  (test-suite
   "test-theme"

   (test-case
    "test-theme1"

    (with-xlsx
     (lambda ()
       (call-with-input-file theme_file
         (lambda (expected)
           (call-with-input-string
            (lists-to-xml (theme))
            (lambda (actual)
              (check-lines? expected actual))))))))
   ))

(run-tests test-theme)
