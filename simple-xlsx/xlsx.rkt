#lang racket

(require "lib/lib.rkt")

(provide (contract-out
          [xlsx% class?]
          [struct sheet
                  (
                  (name string?)
                  (seq exact-nonnegative-integer?)
                  (type symbol?)
                  (typeSeq exact-nonnegative-integer?)
                  (content (or/c data-sheet? line-chart-sheet?))
                  )]
          [struct data-sheet
                  (
                   (rows list?)
                   (width_hash hash?)
                   (color_hash hash?)
                   )]
          [struct line-chart-sheet
                  (
                   (topic string?)
                   (x_data_range data-range?)
                   (y_data_range_list list?)
                   )]
          [struct data-range
                  (
                   (sheet_name string?)
                   (range_str string?)
                   )]
          [convert-range (-> string? string?)]
          [range-length (-> string? exact-nonnegative-integer?)]
          [struct data-serial
                  (
                   (topic string?)
                   (data_range data-range?)
                   )]
          ))

(struct sheet ([name #:mutable] [seq #:mutable] [type #:mutable] [typeSeq #:mutable] [content #:mutable]))

(struct data-sheet ([rows #:mutable] [width_hash #:mutable] [color_hash #:mutable]))
(struct colAttr ([width #:mutable] [back_color #:mutable]))

(struct line-chart-sheet ([topic #:mutable] [x_data_range #:mutable] [y_data_range_list #:mutable]))
(struct data-range ([sheet_name #:mutable] [range_str #:mutable]))
(struct data-serial ([topic #:mutable] [data_range #:mutable]))

(define (convert-range range_str)
  (let loop ([loop_list (regexp-split #rx"-" range_str)]
             [result_str ""])
    (if (not (null? loop_list))
          (let ([items (regexp-match #rx"^([A-Z]+)([0-9]+)$" (car loop_list))])
            (if (string=? result_str "")
                (loop (cdr loop_list) (string-append "$" (second items) "$" (third items) ":"))
                (loop (cdr loop_list) (string-append result_str "$" (second items) "$" (third items)))))
          result_str)))

(define (range-length range_str)
  (let ([numbers (regexp-match* #rx"([0-9]+)" range_str)])
    (add1 (- (string->number (second numbers)) (string->number (first numbers))))))

(define xlsx%
  (class object%
         (super-new)
         
         (field 
          [sheets '()]
          )
         
         (define/public (add-data-sheet sheet_name sheet_data)
           (let* ([sheet_length (length sheets)]
                  [seq (add1 sheet_length)]
                  [type_seq (add1 (length (filter (lambda (rec) (eq? (sheet-type rec) 'data)) sheets)))])
             (set! sheets `(,@sheets
                            ,(sheet
                              sheet_name
                              seq
                              'data
                              type_seq
                              (data-sheet sheet_data (make-hash) (make-hash)))))))
         
         (define/public (get-sheet-by-name sheet_name)
           (let loop ([loop_list sheets])
             (when (not (null? loop_list))
                   (if (string=? sheet_name (sheet-name (car loop_list)))
                       (car loop_list)
                       (loop (cdr loop_list))))))
         
         (define/public (get-range-data sheet_name range_str)
           (let* ([data_sheet (get-sheet-by-name sheet_name)]
                  [col_name (first (regexp-match* #rx"([A-Z]+)" range_str))]
                  [col_number (sub1 (abc->number col_name))]
                  [row_range (regexp-match* #rx"([0-9]+)" range_str)]
                  [row_start_index (string->number (first row_range))]
                  [row_end_index (string->number (second row_range))])
             (let loop ([loop_list (data-sheet-rows (sheet-content data_sheet))]
                        [row_count 1]
                        [result_list '()])
               (if (not (null? loop_list))
                   (if (and (>= row_count row_start_index) (<= row_count row_end_index))
                       (loop (cdr loop_list) (add1 row_count) (cons (list-ref (car loop_list) col_number) result_list))
                       (loop (cdr loop_list) (add1 row_count) result_list))
                   (reverse result_list)))))

         (define/public (add-line-chart-sheet sheet_name topic)
           (let* ([sheet_length (length sheets)]
                  [seq (add1 sheet_length)]
                  [type_seq (add1 (length (filter (lambda (rec) (eq? (sheet-type rec) 'chart)) sheets)))])
             (set! sheets `(,@sheets
                            ,(sheet
                              sheet_name
                              seq
                              'chart
                              type_seq
                              (line-chart-sheet topic (data-range "" "") '()))))))

         (define/public (set-line-chart-x-data! line_chart_sheet_name data_sheet_name data_range)
           (set-line-chart-sheet-x_data_range! (sheet-content (get-sheet-by-name line_chart_sheet_name)) (data-range data_sheet_name data_range)))

         (define/public (add-line-chart-y-data! line_chart_sheet_name y_topic sheet_name data_range)
           (set-line-chart-sheet-y_data_range_list! (sheet-content (get-sheet-by-name line_chart_sheet_name)) `(,@(line-chart-sheet-y_data_range_list (sheet-content (get-sheet-by-name line_chart_sheet_name))) ,(data-serial y_topic (data-range sheet_name data_range)))))
         
         (define/public (sheet-ref sheet_seq)
           (list-ref sheets sheet_seq))
         
         (define/public (set-sheet-col-width! sheet col_range width)
           (hash-set! (data-sheet-width_hash (sheet-content sheet)) col_range width))

         (define/public (set-sheet-col-color! sheet col_range color)
           (hash-set! (data-sheet-color_hash (sheet-content sheet)) col_range color))
         ))

