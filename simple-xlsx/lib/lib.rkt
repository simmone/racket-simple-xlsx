#lang racket

(require file/unzip)
(require file/zip)
(require xml)
(require xml/path)
(require racket/date)
(require rackunit)

(provide (contract-out
          [with-unzip-entry (-> path-string? path-string? (-> path-string? any) any)]
          [with-unzip (-> path-string? (-> path-string? any) any)]
          [xml-get-list (-> symbol? xexpr? list?)]
          [xml-get-attr (-> symbol? string? xexpr? string?)]
          [xml-get (-> symbol? xexpr? string?)]
          [abc->number (-> string? exact-nonnegative-integer?)]
          [abc->range (-> string? pair?)]
          [number->abc (-> number? string?)]
          [number->list (-> number? list?)]
          [format-date (-> date? string?)]
          [format-complete-time (-> date? string?)]
          [format-time (-> number? string?)]
          [value-of-time (-> string? date?)]
          [format-w3cdtf (-> date? string?)]
          [create-sheet-name-list (-> exact-nonnegative-integer? list?)]
          [get-dimension (-> list? string?)]
          [zip-xlsx (-> path-string? path-string? void?)]
          [range-to-cell-hash (-> string? any/c hash?)]
          [combine-hash-in-hash (-> (listof hash?) hash?)]
          [check-lines? (-> input-port? input-port? void?)]
          [prefix-each-line (-> string? string? string?)]
          [date->oa_date_number (-> date? number?)]
          [oa_date_number->date (-> number? date?)]
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

(define (abc->range abc_range)
  (cond
   [(regexp-match #rx"^([0-9]+|[A-Z]+)-([0-9]+|[A-Z]+)$" abc_range)
    (let ([abc_items (regexp-split #rx"-" abc_range)])
      (if (= (length abc_items) 2)
          (let* ([first_item (first abc_items)]
                 [second_item (second abc_items)]
                 [start_index
                  (cond
                   [(regexp-match #rx"^[0-9]+$" first_item)
                    (string->number first_item)]
                   [(regexp-match #rx"^[A-Z]+$" first_item)
                    (abc->number first_item)]
                   [else
                    1])]
                 [end_index
                  (cond
                   [(regexp-match #rx"^[0-9]+$" second_item)
                    (string->number second_item)]
                   [(regexp-match #rx"^[A-Z]+$" second_item)
                    (abc->number second_item)]
                   [else
                    1])])
            (if (<= start_index end_index)
                (cons start_index end_index)
                (cons 1 1)))
          (cons 1 1)))]
   [(regexp-match #rx"^[0-9]+$" abc_range)
    (cons (string->number abc_range) (string->number abc_range))]
   [(regexp-match #rx"^[A-Z]+$" abc_range)
    (cons (abc->number abc_range) (abc->number abc_range))]
   [else
    (cons 1 1)]))

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

(define (get-dimension data_list)
  (let ([rows (length data_list)]
        [cols 0])
    (let loop ([loop_list data_list])
      (when (not (null? loop_list))
            (when (> (length (car loop_list)) cols)
                  (set! cols (length (car loop_list))))
            (loop (cdr loop_list))))

    (string-append (number->abc cols) (number->string rows))))

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

(define (range-to-cell-hash range_str val)
  (let ([flat_map (make-hash)])
    (when (regexp-match #rx"^([A-Z]+)([0-9]+)-([A-Z]+)([0-9]+)$" range_str)
          (let* ([range_items (regexp-match #rx"^([A-Z]+)([0-9]+)-([A-Z]+)([0-9]+)$" range_str)]
                 [start_col_index (abc->number (list-ref range_items 1))]
                 [start_row_index (string->number (list-ref range_items 2))]
                 [end_col_index (abc->number (list-ref range_items 3))]
                 [end_row_index (string->number (list-ref range_items 4))])
            (let range-loop ([loop_col_index start_col_index]
                             [loop_row_index start_row_index])
              (when (and
                     (<= loop_col_index end_col_index)
                     (<= loop_row_index end_row_index))
                    (hash-set! flat_map 
                               (string-append (number->abc loop_col_index) (number->string loop_row_index))
                               val)
                    (cond
                     [(< loop_col_index end_col_index)
                      (range-loop (add1 loop_col_index) loop_row_index)]
                     [(< loop_row_index end_row_index)
                      (range-loop start_col_index (add1 loop_row_index))])))))
    flat_map))

(define (combine-hash-in-hash hash_list)
  (let ([result_map (make-hash)])
    (let outer-hash-loop ([hashes hash_list])
      (when (not (null? hashes))
            (hash-for-each
             (car hashes)
             (lambda (cell_name style_hash)
               (if (hash-has-key? result_map cell_name)
                   (let ([old_hash (hash-copy (hash-ref result_map cell_name))])

                     (hash-for-each
                      style_hash
                      (lambda (ik iv)
                        (hash-set! old_hash ik iv)))
                       (hash-set! result_map cell_name old_hash))
                     (hash-set! result_map cell_name style_hash))))
            (outer-hash-loop (cdr hashes))))
    result_map))

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

(define (date->oa_date_number t_date)
  (let ([epoch (* -1 (find-seconds 0 0 0 30 12 1899))]
        [date_seconds (date->seconds t_date)])
    (inexact->exact (floor (* (/ (+ date_seconds epoch) 86400000) 1000)))))

(define (oa_date_number->date oa_date_number)
  (let* ([epoch (* -1 (find-seconds 0 0 0 30 12 1899))]
         [date_seconds
          (inexact->exact (floor (- (* (/ (floor oa_date_number) 1000) 86400000) epoch)))]
         [actual_date (seconds->date (+ date_seconds (* 24 60 60)))])
    (seconds->date (find-seconds 0 0 0 (date-day actual_date) (date-month actual_date) (date-year actual_date)))))

