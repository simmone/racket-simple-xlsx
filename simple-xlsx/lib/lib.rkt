#lang racket

(require racket/date)
(require rackunit)
(require file/zip)
(require file/unzip)

(require "dimension.rkt")

(provide (contract-out
          [format-w3cdtf (-> date? string?)]
          [zip-xlsx (-> path-string? path-string? void?)]
          [unzip-xlsx (-> path-string? path-string? void?)]
          [check-lines? (-> input-port? input-port? void?)]
          [date->oa_date_number (->* (date?) (boolean?) number?)]
          [oa_date_number->date (->* (number?) (boolean?) date?)]
          [check-data-integrity (-> (listof list?) void?)]
          ))

(define (check-data-integrity data_list)
  (when (equal? data_list '())
        (error "data list is empty"))

  (let loop ([loop_list data_list]
             [child_length -1])
    (when (not (null? loop_list))
          (when (not (list? (car loop_list)))
                (error "data's children is not list type"))

          (when (and
                 (not (= child_length -1))
                 (not (= child_length (length (car loop_list)))))
                (error "data's children's length is not consistent."))

          (loop (cdr loop_list) (length (car loop_list))))))

(define-check (check-lines? expected_port test_port)
  (let* ([expected_lines (port->lines expected_port)]
         [test_lines (port->lines test_port)]
         [test_length (length test_lines)])
    (let loop ([loop_lines expected_lines]
               [line_no 0])
      (when (not (null? loop_lines))
            (cond
             [(>= line_no test_length)
              (fail-check (format "error! line[~a] expected:[~a] actual:null" (add1 line_no) (car loop_lines)))]
             [(not (string=? (car loop_lines) (list-ref test_lines line_no)))
              (fail-check (format "error! line[~a] expected:[~a] actual:[~a]" (add1 line_no) (car loop_lines) (list-ref test_lines line_no)))])
            (loop (cdr loop_lines) (add1 line_no))))))

;; 2014-12-15T13:24:27+08:00
(define (format-w3cdtf the_date)
  (format "~a-~a-~aT~a:~a:~a~a~a:00"
          (~a (date-year the_date) #:min-width 4 #:pad-string "0" #:align 'right)
          (~a (date-month the_date) #:min-width 2 #:pad-string "0" #:align 'right)
          (~a (date-day the_date) #:min-width 2 #:pad-string "0" #:align 'right)
          (~a (date-hour the_date) #:min-width 2 #:pad-string "0" #:align 'right)
          (~a (date-minute the_date) #:min-width 2 #:pad-string "0" #:align 'right)
          (~a (date-second the_date) #:min-width 2 #:pad-string "0" #:align 'right)
          (if (>= (date-time-zone-offset the_date) 0) "+" "-")
          (~a (abs (floor (/ (date-time-zone-offset the_date) 60 60))) #:min-width 2 #:pad-string "0" #:align 'right)))

(define (zip-xlsx zip_file content_dir)
  (let ([pwd #f])
    (dynamic-wind
        (lambda () (set! pwd (current-directory)))
        (lambda ()
          (current-directory content_dir)
          (if (absolute-path? zip_file)
              (zip zip_file "[Content_Types].xml" "_rels" "docProps" "xl")
              (zip (build-path 'up zip_file) "[Content_Types].xml" "_rels" "docProps" "xl")))
        (lambda () (current-directory pwd)))))

(define (unzip-xlsx zip_file content_dir)
  (dynamic-wind
      (lambda () (void))
      (lambda ()
        (call-with-unzip
         zip_file
         (lambda (tmp_dir)
           (delete-directory/files content_dir)
           (copy-directory/files tmp_dir content_dir))))
      (lambda () (void))))

(define (date->oa_date_number t_date [local_time? #t])
  (let ([epoch (* -1 (find-seconds 0 0 0 30 12 1899 local_time?))]
        [date_seconds (date->seconds t_date local_time?)])
    (inexact->exact (floor (* (/ (+ date_seconds epoch) 86400000) 1000)))))

(define (oa_date_number->date oa_date_number [local_time? #t])
  (let* ([epoch (* -1 (find-seconds 0 0 0 30 12 1899 local_time?))]
         [date_seconds
          (inexact->exact (floor (- (* (/ (floor oa_date_number) 1000) 86400000) epoch)))]
         [actual_date (seconds->date (+ date_seconds (* 24 60 60)) local_time?)])
    (seconds->date (find-seconds 0 0 0 (date-day actual_date) (date-month actual_date) (date-year actual_date) local_time?))))
