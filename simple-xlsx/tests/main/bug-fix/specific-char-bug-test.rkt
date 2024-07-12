#lang racket

(require fast-xml
         "../../../main.rkt"
          rackunit/text-ui
          rackunit
          racket/runtime-path)

(define-runtime-path angles_file "angels.xlsx")

(define test-specific-char
  (test-suite
   "test-specific-char"

   (test-case
    "test-specific-char1"

    (dynamic-wind
        (lambda () (void))
        (lambda ()
          (write-xlsx
           angles_file
           (lambda ()
             (add-data-sheet "Sheet1" '(("<test>" "<foo> " " <baz>")
                                        ("< bar>" "< fro >" "<bas >")
                                        ("<maybe" "<< not >>" "show>")))
             ))

          (read-xlsx
           angles_file
           (lambda ()
             (with-sheet-name
              "Sheet1"
              (lambda ()
                (check-equal? (get-row 1) '("<test>" "<foo> " " <baz>"))
                (check-equal? (get-row 2) '("< bar>" "< fro >" "<bas >"))
                (check-equal? (get-row 3) '("<maybe" "<< not >>" "show>"))
                ))
             )))
          (lambda ()
            ;(void)
            (delete-file angles_file)
            )))
   ))

(run-tests test-specific-char)
