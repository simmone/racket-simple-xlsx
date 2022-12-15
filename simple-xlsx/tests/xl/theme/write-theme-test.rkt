#lang racket

(require simple-xml)

(require rackunit/text-ui rackunit)

(require "../../../xlsx/xlsx.rkt")
(require "../../../lib/lib.rkt")

(require"../../../xl/theme/theme.rkt")

(require racket/runtime-path)
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
               (lists->xml (theme))
               (lambda (actual)
                 (check-lines? expected actual))))))
        (lambda ()
          (when (file-exists? theme1_file) (delete-file theme1_file)))))
   ))

(run-tests test-theme)
