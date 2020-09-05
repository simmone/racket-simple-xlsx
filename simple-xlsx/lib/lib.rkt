#lang racket

(require xml)
(require xml/path)
(require racket/date)
(require rackunit)
(require file/zip)

(provide (contract-out
          [xml-get-list (-> symbol? xexpr? list?)]
          [xml-get-attr (-> symbol? string? xexpr? string?)]
          [xml-get (-> symbol? xexpr? string?)]
          [number->list (-> number? list?)]
          [format-date (-> date? string?)]
          [format-complete-time (-> date? string?)]
          [format-time (-> number? string?)]
          [value-of-time (-> string? date?)]
          [format-w3cdtf (-> date? string?)]
          [create-sheet-name-list (-> exact-nonnegative-integer? list?)]
          [zip-xlsx (-> path-string? path-string? void?)]
          [check-lines? (-> input-port? input-port? void?)]
          [prefix-each-line (-> string? string? string?)]
          [date->oa_date_number (->* (date?) (boolean?) number?)]
          [oa_date_number->date (->* (number?) (boolean?) date?)]
          ))

(define-check (check-lines? expected_port test_port)
  (let* ([expected_lines (port->lines expected_port)]
         [test_lines (port->lines test_port)]
         [test_length (sub1 (length test_lines))])
    (let loop ([loop_lines expected_lines]
               [line_no 0])
      (when (not (null? loop_lines))
            (when (or
                   (> line_no test_length)
                   (not (string=? (car loop_lines) (list-ref test_lines line_no))))
                  (fail-check (format "error! line[~a] expected:[~a] actual:[~a]" (add1 line_no) (car loop_lines) (list-ref test_lines line_no))))
            (loop (cdr loop_lines) (add1 line_no))))))

(define (number->list num)
  (let loop ([s_list '()]
             [index 1])
    (if (<= index num)
        `(,@s_list ,@(cons index (loop s_list (add1 index))))
        s_list)))

(define (format-date the_date)
  (format "~a~a~a"
          (~a (date-year the_date) #:min-width 4 #:pad-string "0" #:align 'right)
          (~a (date-month the_date) #:min-width 2 #:pad-string "0" #:align 'right)
          (~a (date-day the_date) #:min-width 2 #:pad-string "0" #:align 'right)))

(define (format-complete-time the_date)
  (format "~a~a~a ~a:~a:~a"
          (~a (date-year the_date) #:min-width 4 #:pad-string "0" #:align 'right)
          (~a (date-month the_date) #:min-width 2 #:pad-string "0" #:align 'right)
          (~a (date-day the_date) #:min-width 2 #:pad-string "0" #:align 'right)
          (~a (date-hour the_date) #:min-width 2 #:pad-string "0" #:align 'right)
          (~a (date-minute the_date) #:min-width 2 #:pad-string "0" #:align 'right)
          (~a (date-second the_date) #:min-width 2 #:pad-string "0" #:align 'right)))

(define (format-time time_s)
  (let* ([hour_s (* time_s 24)]
         [hour (floor hour_s)]
         [hour_sub (- hour_s hour)]
         [minute_s (* hour_sub 60)]
         [minute (floor minute_s)]
         [second_sub (- minute_s minute)]
         [second_s (* second_sub 60)]
         [second (round second_s)])
    (format "~a:~a:~a"
            (~a (inexact->exact hour) #:min-width 2 #:pad-string "0" #:align 'right)
            (~a (inexact->exact minute) #:min-width 2 #:pad-string "0" #:align 'right)
            (~a (inexact->exact second) #:min-width 2 #:pad-string "0" #:align 'right))))

(define (xml-get node-name xml-list)
  (let ([xml_value (se-path* (list node-name) xml-list)])
    (if xml_value
        xml_value
        "")))

(define (xml-get-attr node_name attr_name xml-list)
  (let ([xml_value (se-path* `(,node_name ,(string->keyword attr_name)) xml-list)])
    (if xml_value
        xml_value
        "")))

(define (xml-get-list list-node-name xml-list)
  (se-path*/list (list list-node-name) xml-list))

;; YYYYMMDD HH:MM:SS
(define (value-of-time time_str)
  (let ([year (string->number (substring time_str 0 4))]
        [month (string->number (substring time_str 4 6))]
        [day (string->number (substring time_str 6 8))]
        [hour (string->number (substring time_str 9 11))]
        [minute (string->number (substring time_str 12 14))]
        [second (string->number (substring time_str 15 17))])
    (seconds->date (find-seconds second minute hour day month year))))

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

;; create auto sheet name list: Sheet1, Sheet2, ...
(define (create-sheet-name-list sheet_count)
  (let loop ([sheet_name_list '()]
             [count 1])
    (if (<= count sheet_count)
        (begin
          (set! sheet_name_list `(,@sheet_name_list ,(string-append "Sheet" (number->string count))))
          (loop sheet_name_list (add1 count)))
        sheet_name_list)))

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

(define (prefix-each-line str prefix)
  (with-output-to-string
    (lambda ()
      (let loop ([chars (string->list str)]
                 [is_prefix #t])
          (when (not (null? chars))
                (when (and is_prefix (not (char=? (car chars) #\newline)))
                      (printf "~a" prefix))
                
                (printf "~a" (car chars))

                (if (char=? (car chars) #\newline)
                    (loop (cdr chars) #t)
                    (loop (cdr chars) #f)))))))

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
