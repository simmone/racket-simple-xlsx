#lang racket

(provide (contract-out
          [cell-range? (-> string? boolean?)]
          [cell? (-> string? boolean?)]
          [row_col->cell (-> positive-integer? positive-integer? cell?)]
          [cell->row_col (-> cell? (cons/c positive-integer? positive-integer?))]
          [range->row_col_pair (-> cell-range? (cons/c (cons/c positive-integer? positive-integer?) (cons/c positive-integer? positive-integer?)))]
          [range->capacity (-> cell-range? (cons/c positive-integer? positive-integer?))]
          [capacity->range (->* ((cons/c positive-integer? positive-integer?)) (cell?) cell-range?)]
          [col_abc->number (-> string? positive-integer?)]
          [col_number->abc (-> positive-integer? string?)]
          [to-col-range (-> string? pair?)]
          [to-row-range (-> string? (cons/c positive-integer? positive-integer?))]
          [range->range_xml (-> cell-range? string?)]
          [range_xml->range (-> string? cell-range?)]
          [cell_range->cell_list (-> cell-range? (listof cell?))]
          [get-cell-range-four-sides-cells (-> cell-range? (values (listof cell?) (listof cell?) (listof cell?) (listof cell?)))]
          ))

(define (cell-range? range_str)
  (if(regexp-match #rx"^([A-Z]+)([0-9]+)((-|:)([A-Z]+)([0-9]+))*$" (string-upcase range_str))
     #t
     #f))

(define (cell? cell_str)
  (if (regexp-match #rx"^([A-Z]+)([0-9]+)$" (string-upcase cell_str))
      #t
      #f))

(define (row_col->cell row col)
  (format "~a~a" (col_number->abc col) row))

(define (cell->row_col cell_str)
  (let ([row_col_pair (range->row_col_pair cell_str)])
    (cons (caar row_col_pair) (cdar row_col_pair))))

(define (range->row_col_pair range_str)
  (let ([items_5 (regexp-match #rx"^([A-Z]+)([0-9]+)(-|:)([A-Z]+)([0-9]+)$" (string-upcase range_str))])
    (if items_5
        (let ([start_col_index (col_abc->number (list-ref items_5 1))]
              [start_row_index (string->number (list-ref items_5 2))]
              [end_col_index (col_abc->number (list-ref items_5 4))]
              [end_row_index (string->number (list-ref items_5 5))])
          (if (and
               (>= end_row_index start_row_index)
               (>= end_col_index start_col_index))
              (cons (cons start_row_index start_col_index) (cons end_row_index end_col_index))
              '((1 . 1) . (1 . 1))))
        (let ([items_2 (regexp-match #rx"^([A-Z]+)([0-9]+)$" (string-upcase range_str))])
          (if items_2
              (let ([start_col_index (col_abc->number (list-ref items_2 1))]
                    [start_row_index (string->number (list-ref items_2 2))])
                (cons (cons start_row_index start_col_index) (cons start_row_index start_col_index)))
              '((1 . 1) . (1 . 1)))))))

(define (range->capacity range)
  (let ([row_col_pair (range->row_col_pair range)])
    (cons
     (add1 (- (cadr row_col_pair) (caar row_col_pair)))
     (add1 (- (cddr row_col_pair) (cdar row_col_pair))))))

(define (capacity->range capacity [start_cell? "A1"])
  (let ([start_row_col (cell->row_col start_cell?)])
    (format "~a:~a~a"
            start_cell?
            (col_number->abc (sub1 (+ (cdr start_row_col) (cdr capacity))))
            (sub1 (+ (car start_row_col) (car capacity))))))

(define (range->range_xml range_str)
  (let ([formatted_range_str (string-upcase range_str)])
    (let* ([items (regexp-match #rx"^([A-Z]+)([0-9]+)-([A-Z]+)([0-9]+)$" formatted_range_str)]
           [start_col_name (second items)]
           [start_index (third items)]
           [end_col_name (fourth items)]
           [end_index (fifth items)])
      (format "$~a$~a:$~a$~a" start_col_name start_index end_col_name end_index))))

(define (range_xml->range range_xml_str)
  (let ([formatted_range_str (string-upcase range_xml_str)])
    (let* ([items (regexp-match #rx"^\\$([A-Z]+)\\$([0-9]+):\\$([A-Z]+)\\$([0-9]+)$" formatted_range_str)]
           [start_col_name (second items)]
           [start_index (third items)]
           [end_col_name (fourth items)]
           [end_index (fifth items)])
      (format "~a~a-~a~a" start_col_name start_index end_col_name end_index))))

(define (col_abc->number abc)
  (let ([sum 0])
    (let loop ([char_list (reverse (string->list abc))]
               [base 0])
      (when (not (null? char_list))
            (let* ([alpha (car char_list)]
                   [alpha_int (add1 (- (char->integer alpha) (char->integer #\A)))])
              (set! sum (+ (* alpha_int (expt 26 base)) sum)))
            (loop (cdr char_list) (add1 base))))
    sum))

(define (col_number->abc num)
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

(define (to-col-range col_range)
  (let ([formatted_col_range (string-upcase col_range)])
    (cond
     [(regexp-match #rx"^([0-9]+|[A-Z]+)-([0-9]+|[A-Z]+)$" formatted_col_range)
      (let ([abc_items (regexp-split #rx"-" formatted_col_range)])
        (if (= (length abc_items) 2)
            (let* ([first_item (first abc_items)]
                   [second_item (second abc_items)]
                   [start_index
                    (cond
                     [(regexp-match #rx"^[0-9]+$" first_item)
                      (string->number first_item)]
                     [(regexp-match #rx"^[A-Z]+$" first_item)
                      (col_abc->number first_item)]
                     [else
                      1])]
                   [end_index
                    (cond
                     [(regexp-match #rx"^[0-9]+$" second_item)
                      (string->number second_item)]
                     [(regexp-match #rx"^[A-Z]+$" second_item)
                      (col_abc->number second_item)]
                     [else
                      1])])
              (if (<= start_index end_index)
                  (cons start_index end_index)
                  (cons 1 1)))
            (cons 1 1)))]
     [(regexp-match #rx"^[0-9]+$" formatted_col_range)
      (cons (string->number formatted_col_range) (string->number formatted_col_range))]
     [(regexp-match #rx"^[A-Z]+$" formatted_col_range)
      (cons (col_abc->number formatted_col_range) (col_abc->number formatted_col_range))]
     [else
      (cons 1 1)])))

(define (to-row-range row_range_str)
  (cond
   [(regexp-match #rx"^[0-9]+$" row_range_str)
    (let* ([items (regexp-match #rx"^([0-9]+)$" row_range_str)]
           [start_row_index (second items)]
           [end_row_index start_row_index])
      (cons (string->number start_row_index) (string->number end_row_index)))]
   [(regexp-match #rx"^[0-9]+-[0-9]+$" row_range_str)
    (let* ([items (regexp-match #rx"^([0-9]+)-([0-9]+)$" row_range_str)]
           [start_row_index (string->number (second items))]
           [end_row_index (string->number (third items))])
      (if (> start_row_index end_row_index)
          '(1 . 1)
          (cons start_row_index end_row_index)))]
   [else
    '(1 . 1)]))

(define (cell_range->cell_list range_str)
  (let* ([range_pair (range->row_col_pair range_str)]
         [start_col_index (cdar range_pair)]
         [start_row_index (caar range_pair)]
         [end_col_index (cddr range_pair)]
         [end_row_index (cadr range_pair)])
    (let loop ([loop_col_index start_col_index]
               [loop_row_index start_row_index]
               [cell_list '()])
      (if (<= loop_row_index end_row_index)
          (if (<= loop_col_index end_col_index)
              (loop (add1 loop_col_index) loop_row_index
                    (cons
                     (format "~a~a" (col_number->abc loop_col_index) (number->string loop_row_index))
                     cell_list))
              (loop start_col_index (add1 loop_row_index) cell_list))
          (reverse cell_list)))))

(define (get-cell-range-four-sides-cells cell_range_str)
  (let* ([row_col_pair (range->row_col_pair cell_range_str)]
         [start_row (caar row_col_pair)]
         [start_col (cdar row_col_pair)]
         [end_row (cadr row_col_pair)]
         [end_col (cddr row_col_pair)])

    (let loop ([cells (cell_range->cell_list cell_range_str)]
               [top_cells '()]
               [bottom_cells '()]
               [left_cells '()]
               [right_cells '()])
      (if (not (null? cells))
          (let ([row_col (cell->row_col (car cells))]
                [loop_top_cells top_cells]
                [loop_bottom_cells bottom_cells]
                [loop_left_cells left_cells]
                [loop_right_cells right_cells])

            (when (= (car row_col) start_row) (set! loop_top_cells (cons (car cells) loop_top_cells)))
            (when (= (car row_col) end_row)  (set! loop_bottom_cells (cons (car cells) loop_bottom_cells)))
            (when (= (cdr row_col) start_col) (set! loop_left_cells (cons (car cells) loop_left_cells)))
            (when (= (cdr row_col) end_col) (set! loop_right_cells (cons (car cells) loop_right_cells)))

            (loop (cdr cells) loop_top_cells loop_bottom_cells loop_left_cells loop_right_cells))
          (values
           (reverse top_cells) (reverse bottom_cells) (reverse left_cells) (reverse right_cells))))))
