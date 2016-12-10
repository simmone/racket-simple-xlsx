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
                  (content (or/c data-sheet? chart-sheet?))
                  )]
          [struct data-sheet
                  (
                   (rows list?)
                   (width_hash hash?)
                   (color_hash hash?)
                   )]
          [struct chart-sheet
                  (
                   (chart_type symbol?)
                   (topic string?)
                   (x_topic string?)
                   (x_data_range data-range?)
                   (y_data_range_list list?)
                   )]
          [check-range (-> string? boolean?)]
          [check-col-range (-> string? boolean?)]
          [check-cell-range (-> string? boolean?)]
          [check-data-range-valid (-> (is-a?/c xlsx%) string? string? boolean?)]
          [check-data-list (-> any/c boolean?)]
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

(struct chart-sheet ([chart_type #:mutable] [topic #:mutable] [x_topic #:mutable] [x_data_range #:mutable] [y_data_range_list #:mutable]))
(struct data-range ([sheet_name #:mutable] [range_str #:mutable]))
(struct data-serial ([topic #:mutable] [data_range #:mutable]))

(define (check-data-list data_list)
  (when (not (list? data_list))
        (error "data is not list type"))
  
  (when (equal? data_list '())
        (error "data has no children list"))
  
  (let loop ([loop_list data_list]
             [child_length -1])
    (when (not (null? loop_list))
          (when (not (list? (car loop_list)))
                (error "data's children is not list type"))
          
          (when (and
                 (not (= child_length -1))
                 (not (= child_length (length (car loop_list)))))
                (error "data's children's length is not consistent."))
          
          (loop (cdr loop_list) (length (car loop_list)))))
  #t)

(define (check-range range_str)
  (if (regexp-match #rx"^[A-Z]+[0-9]+-[A-Z]+[0-9]+$" range_str)
      (let* ([items (regexp-match #rx"^([A-Z]+)([0-9]+)-([A-Z]+)([0-9]+)$" range_str)]
             [start_col_name (second items)]
             [start_row_index (third items)]
             [end_col_name (fourth items)]
             [end_row_index (fifth items)])

        (if (string=? start_col_name end_col_name)
            (if (> (string->number start_row_index) (string->number end_row_index))
                (error (format "range's direction is vertical, index is invalid.[~a][~a]" start_row_index end_row_index))
                #t)
            (if (string=? start_row_index end_row_index)
                (if (> (abc->number start_col_name) (abc->number end_col_name))
                    (error (format "range's direction is horizontal, col name is invalid.[~a][~a]" start_col_name end_col_name))
                    #t)
                (error (format "range's direction confused. should like A1-A20 or A2-Z2, but get ~a" range_str)))))
      (error (format "range format should like A1-A20 or A2-Z2, but get ~a" range_str))))

(define (check-col-range col_range_str)
  (cond
   [(regexp-match #rx"^[A-Z]+-[A-Z]+$" col_range_str)
    (let* ([items (regexp-match #rx"^([A-Z]+)-([A-Z]+)$" col_range_str)]
           [start_col_name (second items)]
           [end_col_name (third items)])
      (if (string>? start_col_name end_col_name)
          (error (format "col name should from small to big[~a]" col_range_str))
          #t))]
   [(regexp-match #rx"^[0-9]+-[0-9]+$" col_range_str)
    (let* ([items (regexp-match #rx"^([0-9]+)-([0-9]+)$" col_range_str)]
           [start_col_index (second items)]
           [end_col_index (third items)])
      (if (string>? start_col_index end_col_index)
          (error (format "col index should from small to big[~a]" col_range_str))
          #t))]
   [else
    (error (format "invalid col range! should be like this: A-Z or 1-10[~a]" col_range_str))]
   ))

(define (check-cell-range cell_range_str)
  (if (regexp-match #rx"^[A-Z]+[0-9]+-[A-Z]+[0-9]+$" cell_range_str)
      (let* ([items (regexp-match #rx"^([A-Z]+)([0-9]+)-([A-Z]+)([0-9]+)$" cell_range_str)]
             [start_col_name (second items)]
             [start_col_index (string->number (third items))]
             [end_col_name (fourth items)]
             [end_col_index (string->number (fifth items))])
        (cond
         [(string<? end_col_name start_col_name)
          (error (format "col name should from small to big[~a]" cell_range_str))]
         [(< end_col_index start_col_index)
          (error (format "col index should from small to big[~a]" cell_range_str))]
         [else
          #t]))
      (error (format "invalid cell range! should be like this: A1-B2[~a]" cell_range_str))))

(define (convert-range range_str)
  (when (check-range range_str)
        (let* ([items (regexp-match #rx"^([A-Z]+)([0-9]+)-([A-Z]+)([0-9]+)$" range_str)]
               [start_col_name (second items)]
               [start_index (third items)]
               [end_col_name (fourth items)]
               [end_index (fifth items)])
          (string-append "$" start_col_name "$" start_index ":$" end_col_name "$" end_index))))

(define (range-length range_str)
  (let* ([items (regexp-match #rx"^([A-Z]+)([0-9]+)-([A-Z]+)([0-9]+)$" range_str)]
         [start_col_name (second items)]
         [start_row_index (third items)]
         [end_col_name (fourth items)]
         [end_row_index (fifth items)])

         (if (string=? start_col_name end_col_name)
             (add1 (- (string->number end_row_index) (string->number start_row_index)))
             (add1 (- (abc->number end_col_name) (abc->number start_col_name))))))

(define (check-data-range-valid xlsx sheet_name range_str)
  (when (check-range range_str)
        (let* ([rows (data-sheet-rows (sheet-content (send xlsx get-sheet-by-name sheet_name)))]
               [first_row (first rows)]
               [col_name (first (regexp-match* #rx"([A-Z]+)" range_str))]
               [col_number (sub1 (abc->number col_name))]
               [row_range (regexp-match* #rx"([0-9]+)" range_str)]
               [row_start_index (string->number (first row_range))]
               [row_end_index (string->number (second row_range))])
          (cond
           [(< (length first_row) (add1 col_number))
            (error (format "no such column[~a]" col_name))]
           [(> row_end_index (length rows))
            (error (format "end index beyond data range[~a]" row_end_index))]
           [else
            #t]))))

(define xlsx%
  (class object%
         (super-new)
         
         (field 
          [sheets '()]
          [sheet_name_map (make-hash)]
          [string_item_map (make-hash)]
          )
         
         (define/public (add-data-sheet #:sheet_name sheet_name #:sheet_data sheet_data)
           (when (check-data-list sheet_data)
                 (if (not (hash-has-key? sheet_name_map sheet_name))
                     (let* ([sheet_length (length sheets)]
                            [seq (add1 sheet_length)]
                            [type_seq (add1 (length (filter (lambda (rec) (eq? (sheet-type rec) 'data)) sheets)))])
                       
                       (let loop ([loop_list sheet_data])
                         (when (not (null? loop_list))
                               (let inner-loop ([inner_loop_list (car loop_list)])
                                 (when (not (null? inner_loop_list))
                                       (when (string? (car inner_loop_list))
                                             (hash-set! string_item_map (car inner_loop_list) ""))
                                       (inner-loop (cdr inner_loop_list))))
                               (loop (cdr loop_list))))
                       
                       (set! sheets `(,@sheets
                                      ,(sheet
                                        sheet_name
                                        seq
                                        'data
                                        type_seq
                                        (data-sheet sheet_data (make-hash) (make-hash)))))
                       (hash-set! sheet_name_map sheet_name (sub1 seq)))
                     (error (format "duplicate sheet name[~a]" sheet_name)))))
         
         (define/public (get-sheet-by-name sheet_name)
           (if (hash-has-key? sheet_name_map sheet_name)
               (list-ref sheets (hash-ref sheet_name_map sheet_name))
               (error (format "no such sheet[~a]" sheet_name))))
         
         (define/public (get-range-data sheet_name range_str)
           (when (check-data-range-valid this sheet_name range_str)
                 (let* ([data_sheet (get-sheet-by-name sheet_name)]
                        [col_range (regexp-match* #rx"([A-Z]+)" range_str)]
                        [start_col_name (first col_range)]
                        [col_start_index (sub1 (abc->number start_col_name))]
                        [end_col_name (second col_range)]
                        [col_end_index (sub1 (abc->number end_col_name))]
                        [row_range (regexp-match* #rx"([0-9]+)" range_str)]
                        [row_start_index (string->number (first row_range))]
                        [row_end_index (string->number (second row_range))]
                        [direction (if (string=? start_col_name end_col_name) 'vertical 'horizontal)])
                   (if (eq? direction 'vertical)
                       (let loop ([loop_list (data-sheet-rows (sheet-content data_sheet))]
                                  [row_count 1]
                                  [result_list '()])
                         (if (not (null? loop_list))
                             (if (and (>= row_count row_start_index) (<= row_count row_end_index))
                                 (loop (cdr loop_list) (add1 row_count) 
                                       (cons (list-ref (car loop_list) col_start_index) result_list))
                                 (loop (cdr loop_list) (add1 row_count) result_list))
                             (reverse result_list)))
                       (let loop ([loop_list (list-ref (data-sheet-rows (sheet-content data_sheet)) (sub1 row_start_index))]
                                  [col_index 0]
                                  [result_list '()])
                         (if (not (null? loop_list))
                             (if (and (>= col_index col_start_index) (<= col_index col_end_index))
                                 (loop (cdr loop_list) (add1 col_index) (cons (car loop_list) result_list))
                                 (loop (cdr loop_list) (add1 col_index) result_list))
                             (reverse result_list)))))))

         (define/public (add-chart-sheet #:sheet_name sheet_name #:chart_type [chart_type 'line] #:topic [topic ""] #:x_topic [x_topic ""])
           (if (not (hash-has-key? sheet_name_map sheet_name))
               (let* ([sheet_length (length sheets)]
                      [seq (add1 sheet_length)]
                      [type_seq (add1 (length (filter (lambda (rec) (eq? (sheet-type rec) 'chart)) sheets)))])
                 (set! sheets `(,@sheets
                                ,(sheet
                                  sheet_name
                                  seq
                                  'chart
                                  type_seq
                                  (chart-sheet chart_type topic x_topic (data-range "" "") '()))))
                 (hash-set! sheet_name_map sheet_name (sub1 seq)))
               (error (format "duplicate sheet name[~a]" sheet_name))))
                 
         (define/public (set-chart-x-data! #:sheet_name sheet_name #:data_sheet_name data_sheet_name #:data_range data_range)
           (when (check-data-range-valid this data_sheet_name data_range)
                 (set-chart-sheet-x_data_range! (sheet-content (get-sheet-by-name sheet_name)) (data-range data_sheet_name data_range))))

         (define/public 
           (add-chart-serial! #:sheet_name sheet_name #:data_sheet_name data_sheet_name #:data_range data_range #:y_topic [y_topic ""])
           (when (check-data-range-valid this data_sheet_name data_range)
                 (set-chart-sheet-y_data_range_list! (sheet-content (get-sheet-by-name sheet_name)) `(,@(chart-sheet-y_data_range_list (sheet-content (get-sheet-by-name sheet_name))) ,(data-serial y_topic (data-range data_sheet_name data_range))))))

         (define/public (sheet-ref sheet_seq)
           (list-ref sheets sheet_seq))
         
         (define/public (set-data-sheet-col-width! #:sheet_name sheet_name #:col_range col_range #:width width)
           (when (check-col-range col_range)
                 (hash-set! (data-sheet-width_hash (sheet-content (get-sheet-by-name sheet_name))) col_range width)))

         (define/public (set-data-sheet-cell-color! #:sheet_name sheet_name #:cell_range cell_range #:color color)
           (when (check-cell-range cell_range)
                 (hash-set! (data-sheet-color_hash (sheet-content (get-sheet-by-name sheet_name))) cell_range color)))

         (define/public (get-string-item-list)
           (sort (hash-keys string_item_map) string<?))

         (define/public (get-string-index-map)
           (let ([string_index_map (make-hash)])
             (let loop ([loop_list (get-string-item-list)]
                        [index 0])
               (when (not (null? loop_list))
                     (hash-set! string_index_map (car loop_list) index)
                     (loop (cdr loop_list) (add1 index))))
             string_index_map))

         (define/public (get-color-list)
           (let ([style_list '()]
                 [tmp_hash (make-hash)])
             (let loop ([loop_list sheets])
               (when (not (null? loop_list))
                     (when (eq? (sheet-type (car loop_list)) 'data)
                           (let ([color_hash (data-sheet-color_hash (sheet-content (car loop_list)))])
                             (hash-for-each
                              color_hash
                              (lambda (range_str color_str)
                                (hash-set! tmp_hash color_str "")))))
                     (loop (cdr loop_list))))
             
             (sort (hash-keys tmp_hash) string<?)))

         ))

