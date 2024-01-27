#lang racket

(require rackunit/text-ui rackunit)
(require racket/date)

(require"../../lib/lib.rkt")

(require racket/runtime-path)
(define-runtime-path zip_xlsx_temp_directory "test-directory")
(define-runtime-path content_type_file (build-path "test-directory" "[Content_Types].xml"))
(define-runtime-path rels_directory (build-path "test-directory" "_rels"))
(define-runtime-path doc_props_directory (build-path "test-directory" "docProps"))
(define-runtime-path xl_directory (build-path "test-directory" "xl"))
(define-runtime-path zip_xlsx_file "test.xlsx")

(define test-lib
  (test-suite
   "test-lib"

   (test-case
    "test-maintain-sheet-data-consistency"

    (check-exn exn:fail? (lambda () (maintain-sheet-data-consistency '() "")))
    (check-exn exn:fail? (lambda () (maintain-sheet-data-consistency '((1) 4) "")))

    (check-equal? (maintain-sheet-data-consistency '((1) (1 2)) "")
                 '((1 "") (1 2)))

    (check-equal? (maintain-sheet-data-consistency '((1) (1 2)) 0)
                 '((1 0) (1 2)))

    (check-equal? (maintain-sheet-data-consistency '((1 2) (3 4)) 0)
                 '((1 2) (3 4)))
    )

   (test-case
    "test-check-lines1"
    (call-with-input-string
     "abc"
     (lambda (expected_port)
       (call-with-input-string
        "abc"
        (lambda (test_port)
          (check-lines? expected_port test_port))))))

   (test-case
    "test-check-lines2"
    (call-with-input-string
     "abc\n11"
     (lambda (expected_port)
       (call-with-input-string
        "abc\n11"
        (lambda (test_port)
          (check-lines? expected_port test_port))))))

   (test-case
    "test-check-lines3"
    (call-with-input-string
     " a\n\n b\n\nc\n"
     (lambda (expected_port)
       (call-with-input-string
        " a\n\n b\n\nc\n"
        (lambda (test_port)
          (check-lines? expected_port test_port))))))

   (test-case
    "test-format-w3cdtf"
    (check-equal? (format-w3cdtf (date* 44 17 13 2 1 2015 5 1 #f 28800 996159076 "CST")) "2015-01-02T13:17:44+08:00"))

   (test-case
    "test-date->oadate"

    (check-equal? (date->oa_date_number (seconds->date (find-seconds 0 0 0 17 9 2018 #f)) #f) 43360)

    (check-equal? (date->oa_date_number (seconds->date (find-seconds 0 0 0 16 9 2018 #f)) #f) 43359)
    )

   (test-case
    "test-oadate->date"

    (check-equal? (oa_date_number->date 43360 #f) (seconds->date (find-seconds 0 0 0 18 9 2018 #f)))

    (check-equal? (oa_date_number->date 43359 #f) (seconds->date (find-seconds 0 0 0 17 9 2018 #f)))

    (check-equal? (oa_date_number->date 43359.1212121 #f) (seconds->date (find-seconds 0 0 0 17 9 2018 #f)))
    )

   (test-case
    "test-zip-and-unzip"

    (dynamic-wind
        (lambda ()
          (make-directory zip_xlsx_temp_directory)
          (write-to-file "" content_type_file)
          (make-directory rels_directory)
          (make-directory xl_directory)
          (make-directory doc_props_directory)
          (zip-xlsx zip_xlsx_file zip_xlsx_temp_directory)
          )
        (lambda ()
          (unzip-xlsx zip_xlsx_file zip_xlsx_temp_directory)
          (check-true (file-exists? content_type_file))
          (check-true (directory-exists? rels_directory))
          (check-true (directory-exists? doc_props_directory))
          (check-true (directory-exists? xl_directory)))
        (lambda ()
          (delete-file zip_xlsx_file)
          (delete-directory/files zip_xlsx_temp_directory)
          )))
   ))

(run-tests test-lib)
