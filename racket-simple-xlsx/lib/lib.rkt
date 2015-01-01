#lang racket

(require file/unzip)
(require xml)
(require xml/path)
(require racket/date)

(provide (contract-out
          [with-unzip-entry (-> path-string? path-string? (-> path-string? any) any)]
          [with-unzip (-> path-string? (-> path-string? any) any)]
          [xml-get-list (-> symbol? xexpr? list?)]
          [xml-get-attr (-> symbol? string? xexpr? string?)]
          [xml-get (-> symbol? xexpr? string?)]
          [abc->number (-> string? number?)]
          [number->abc (-> number? string?)]
          [number->list (-> number? list?)]
          [format-date (-> date? string?)]
          [format-time (-> number? string?)]
          ))

(define (format-date the_date)
  (format "~a-~a-~a" (date-year the_date) (date-month the_date) (date-day the_date)))

(define (string-fill str char len #:direction [direction 'left])
  (let* ([str_len (string-length str)]
         [distance (- len str_len)])

    (if (> distance 0)
        (with-output-to-string 
          (lambda ()
            
            (when (eq? direction 'right)
              (printf str))

            (let ([count 0])
              (letrec ([recur
                        (lambda ()
                          (set! count (add1 count))

                          (when (<= count distance)
                            (printf "~a" char)
                            (recur)))])
                (recur)))

            (when (eq? direction 'left)
              (printf str))))
        str)))

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
            (string-fill (number->string (inexact->exact hour)) #\0 2)
            (string-fill (number->string (inexact->exact minute)) #\0 2)
            (string-fill (number->string (inexact->exact second)) #\0 2))))

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

(define (with-unzip-entry zip_file entry_file do_proc)
  (let ([temp_dir #f])
    (dynamic-wind
        (lambda ()
          (set! temp_dir (make-temporary-file "ziptmp~a" 'directory ".")))
        (lambda ()
          (let ([directory_entries (read-zip-directory zip_file)])
            (unzip-entry 
             zip_file
             directory_entries 
             (path->zip-path entry_file) 
             (make-filesystem-entry-reader #:dest temp_dir #:exists 'replace))
            (do_proc (build-path temp_dir entry_file))))
        (lambda ()
          (delete-directory/files temp_dir)))))

(define (with-unzip zip_file do_proc)
  (let ([temp_dir #f])
    (dynamic-wind
        (lambda ()
          (set! temp_dir (make-temporary-file "ziptmp~a" 'directory ".")))
        (lambda ()
          (unzip zip_file (make-filesystem-entry-reader #:dest temp_dir #:exists 'replace))
          (do_proc temp_dir))
        (lambda ()
;          (void)))))
          (delete-directory/files temp_dir)))))

(define (abc->number abc)
  (let ([sum 0])
    (let loop ([char_list (reverse (string->list abc))]
               [base 0])
      (when (not (null? char_list))
            (let* ([alpha (car char_list)]
                   [alpha_int (add1 (- (char->integer alpha) (char->integer #\A)))])
              (set! sum (+ (* alpha_int (expt 26 base)) sum)))
            (loop (cdr char_list) (add1 base))))
    sum))

(define (number->abc num)
  (let ([abc ""])
    (let loop ([loop_num num])
      (if (> loop_num 26)
          (let-values ([(quo remain) (quotient/remainder loop_num 26)])
            (if (= remain 0)
                (begin
                  (set! abc (string-append "Z" abc))
                  (loop (sub1 quo)))
                (begin
                  (set! abc (string-append (string (integer->char (+ 64 remain))) abc))
                  (loop quo))))
            (set! abc (string-append (string (integer->char (+ 64 loop_num))) abc))))
    abc))

(define (number->list num)
  (let loop ([s_list '()]
             [index 1])
    (if (<= index num)
        `(,@s_list ,@(cons index (loop s_list (add1 index))))
        s_list)))
    
