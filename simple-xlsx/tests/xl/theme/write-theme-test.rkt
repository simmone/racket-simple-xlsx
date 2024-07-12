#lang racket

(require fast-xml
         rackunit/text-ui
         rackunit
         "../../../xlsx/xlsx.rkt"
         "../../../lib/lib.rkt"
         "../../../xl/theme/theme.rkt"
         racket/runtime-path)

(define-runtime-path theme1_file "theme1.xml")

(define test-theme
  (test-suite
   "test-theme"

   (test-case
    "test-write-theme"

    (dynamic-wind
        (lambda ()
          (write-theme (apply build-path (drop-right (explode-path theme1_file) 1))))
        (lambda ()
          (call-with-input-file theme1_file
            (lambda (expected)
              (call-with-input-string
               (lists-to-xml (theme))
               (lambda (actual)
                 (check-lines? expected actual))))))
        (lambda ()
          (when (file-exists? theme1_file) (delete-file theme1_file)))))
   ))

(run-tests test-theme)
