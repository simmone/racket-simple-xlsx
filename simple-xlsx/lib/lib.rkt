#lang racket

(require file/unzip)
(require file/zip)
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
          [abc->range (-> string? pair?)]
          [number->abc (-> number? string?)]
          [get-range-ref (-> hash? number? (or/c any/c #f))]
          [number->list (-> number? list?)]
          [format-date (-> date? string?)]
          [format-complete-time (-> date? string?)]
          [format-time (-> number? string?)]
          [value-of-time (-> string? date?)]
          [format-w3cdtf (-> date? string?)]
          [create-sheet-name-list (-> exact-nonnegative-integer? list?)]
          [get-dimension (-> list? string?)]
          [get-string-index (-> list? (values list? hash?))]
          [zip-xlsx (-> path-string? path-string? void?)]
          [struct col-attr ((width number?) (color string?))]
          [cx-round (-> number? integer? number?)]          
          ))

(define (cx-round num precise)
  (let ([factor (expt 10 precise)])
    (/ (round (* num factor)) factor)))

;; represent a column's attibutes, if not want set a specific attr, set it to #f
;; example: (col-attr #f "red") means only set color
(struct col-attr (width color) #:transparent)

(define (format-date the_date)
  (format "~a-~a-~a" (date-year the_date) (date-month the_date) (date-day the_date)))

(define (format-complete-time the_date)
  (format "~a~a~a ~a:~a:~a" 
          (string-fill (number->string (date-year the_date)) #\0 4)
          (string-fill (number->string (date-month the_date)) #\0 2)
          (string-fill (number->string (date-day the_date)) #\0 2)
          (string-fill (number->string (date-hour the_date)) #\0 2)
          (string-fill (number->string (date-minute the_date)) #\0 2)
          (string-fill (number->string (date-second the_date)) #\0 2)))

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

(define (get-range-ref range_hash index)
  (let ([result #f])
    (let loop ([data_list (hash->list range_hash)])
      (when (not (null? data_list))
            (let ([range (caar data_list)]
                  [value (cdar data_list)])
       (if (and (>= index (car range)) (<= index (cdr range)))
           (set! result value)
           (loop (cdr data_list))))))
    result))

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

;; "3" -> "03" or "30"
(define (string-fill str char len #:direction [direction 'left])
  (let* ([str_len (string-length str)]
         [distance (- len str_len)])

    (if (> distance 0)
        (with-output-to-string 
          (lambda ()
            (when (eq? direction 'right)
              (printf str))

            (let loop ([count 1])
              (when (<= count distance)
                    (printf "~a" char)
                    (loop (add1 count))))

            (when (eq? direction 'left)
                  (printf str))))
        str)))

;; 2014-12-15T13:24:27+08:00
(define (format-w3cdtf the_date)
  (format "~a-~a-~aT~a:~a:~a~a~a:00" 
          (date-year the_date) 
          (string-fill (number->string (date-month the_date)) #\0 2)
          (string-fill (number->string (date-day the_date)) #\0 2)
          (string-fill (number->string (date-hour the_date)) #\0 2)
          (string-fill (number->string (date-minute the_date)) #\0 2)
          (string-fill (number->string (date-second the_date)) #\0 2)
          (if (>= (date-time-zone-offset the_date) 0) "+" "-")
          (string-fill (number->string (abs (floor (/ (date-time-zone-offset the_date) 60 60)))) #\0 2)))

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
