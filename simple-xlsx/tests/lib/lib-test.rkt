#lang racket

(require rackunit/text-ui)
(require racket/date)

(require rackunit "../../lib/lib.rkt")

(define test-lib
  (test-suite
   "test-lib"
   
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
    "test-number->list"
    (check-equal? (number->list 1) '(1))
    (check-equal? (number->list 2) '(1 2))
    (check-equal? (number->list 5) '(1 2 3 4 5)))
   
   (test-case
    "test-format-hour"
    (check-equal? (format-time 0.5560763888888889) "13:20:45")
    (check-equal? (format-time 0.2560763888888889) "06:08:45")
    (check-equal? (format-time 0.5) "12:00:00"))
   
   (test-case
    "test-value-of-time"
    (check-equal? (format-complete-time (value-of-time "20141231 12:31:23")) "20141231 12:31:23")
    (check-equal? (format-complete-time (value-of-time "20140105 12:31:23")) "20140105 12:31:23"))

   (test-case
    "test-format-w3cdtf"
    (check-equal? (format-w3cdtf (date* 44 17 13 2 1 2015 5 1 #f 28800 996159076 "CST")) "2015-01-02T13:17:44+08:00"))
   
   (test-case
    "test-create-sheet-name-list"
    (check-equal? (create-sheet-name-list 5) '("Sheet1" "Sheet2" "Sheet3" "Sheet4" "Sheet5"))
    (check-equal? (create-sheet-name-list 1) '("Sheet1")))

   (test-case
    "test-prefix-each-line"
    
    (let ([str "kkd\nskdfk\n\nksjdkf\n\n"])
      (check-equal? (prefix-each-line str "  ")
                    "  kkd\n  skdfk\n\n  ksjdkf\n\n"))

    (let ([str "kkd\nskdfk\n\nksjdkf"])
      (check-equal? (prefix-each-line str "  ")
                    "  kkd\n  skdfk\n\n  ksjdkf"))

    (let ([str "kkd\nskdfk\n\nksjdkf\n"])
      (check-equal? (prefix-each-line str "  ")
                    "  kkd\n  skdfk\n\n  ksjdkf\n"))

    (let ([str ""])
      (check-equal? (prefix-each-line str "  ") ""))

    (let ([str "\n"])
      (check-equal? (prefix-each-line str "  ") "\n"))

    (let ([str "\n\n\n"])
      (check-equal? (prefix-each-line str "  ") "\n\n\n"))

    (let ([str "\nsskd\n\n"])
      (check-equal? (prefix-each-line str "  ") "\n  sskd\n\n"))
    )

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
    
  ))

(run-tests test-lib)
