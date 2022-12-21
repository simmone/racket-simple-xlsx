#lang racket

(require "../xlsx/xlsx.rkt")
(require "../sheet/sheet.rkt")
(require "dimension.rkt")

(provide (contract-out
          [get-sheet-dimension (-> string?)]

          [get-rows-count (-> natural?)]
          [get-sheet-ref-rows-count (-> natural? natural?)]
          [get-sheet-name-rows-count (-> string? natural?)]
          [get-sheet-*name*-rows-count (-> string? natural?)]

          [get-cols-count (-> natural?)]
          [get-sheet-ref-cols-count (-> natural? natural?)]
          [get-sheet-name-cols-count (-> string? natural?)]
          [get-sheet-*name*-cols-count (-> string? natural?)]

          [get-row-cells (-> natural? (listof string?))]
          [get-sheet-ref-row-cells (-> natural? natural? (listof string?))]
          [get-sheet-name-row-cells (-> string? natural? (listof string?))]
          [get-sheet-*name*-row-cells (-> string? natural? (listof string?))]

          [get-col-cells (-> (or/c natural? string?) (listof string?))]
          [get-sheet-ref-col-cells (-> natural? (or/c natural? string?) (listof string?))]
          [get-sheet-name-col-cells (-> string? (or/c natural? string?) (listof string?))]
          [get-sheet-*name*-col-cells (-> string? (or/c natural? string?) (listof string?))]

          [get-cell (-> string? cell-value?)]
          [get-sheet-ref-cell (-> natural? string? cell-value?)]
          [get-sheet-name-cell (-> string? string? cell-value?)]
          [get-sheet-*name*-cell (-> string? string? cell-value?)]

          [set-cell! (-> string? cell-value? void?)]
          [set-sheet-ref-cell! (-> natural? string? cell-value? void?)]
          [set-sheet-name-cell! (-> string? string? cell-value? void?)]
          [set-sheet-*name*-cell! (-> string? string? cell-value? void?)]

          [get-row (-> natural? (listof cell-value?))]
          [get-sheet-ref-row (-> natural? natural? (listof cell-value?))]
          [get-sheet-name-row (-> string? natural? (listof cell-value?))]
          [get-sheet-*name*-row (-> string? natural? (listof cell-value?))]

          [set-row! (-> natural? (listof cell-value?) void?)]
          [set-sheet-ref-row! (-> natural? natural? (listof cell-value?) void?)]
          [set-sheet-name-row! (-> string? natural? (listof cell-value?) void?)]
          [set-sheet-*name*-row! (-> string? natural? (listof cell-value?) void?)]

          [get-rows (-> (listof (listof cell-value?)))]
          [get-sheet-ref-rows (-> natural? (listof (listof cell-value?)))]
          [get-sheet-name-rows (-> string? (listof (listof cell-value?)))]
          [get-sheet-*name*-rows (-> string? (listof (listof cell-value?)))]

          [set-rows! (-> (listof (listof cell-value?)) void?)]
          [set-sheet-ref-rows! (-> natural? (listof (listof cell-value?)) void?)]
          [set-sheet-name-rows! (-> string? (listof (listof cell-value?)) void?)]
          [set-sheet-*name*-rows! (-> string? (listof (listof cell-value?)) void?)]

          [get-col (-> (or/c natural? string?) (listof cell-value?))]
          [get-sheet-ref-col (-> natural? (or/c natural? string?) (listof cell-value?))]
          [get-sheet-name-col (-> string? (or/c natural? string?) (listof cell-value?))]
          [get-sheet-*name*-col (-> string? (or/c natural? string?) (listof cell-value?))]

          [set-col! (-> (or/c natural? string?) (listof cell-value?) void?)]
          [set-sheet-ref-col! (-> natural? (or/c natural? string?) (listof cell-value?) void?)]
          [set-sheet-name-col! (-> string? (or/c natural? string?) (listof cell-value?) void?)]
          [set-sheet-*name*-col! (-> string? (or/c natural? string?) (listof cell-value?) void?)]

          [get-cols (-> (listof (listof cell-value?)))]
          [get-sheet-ref-cols (-> natural? (listof (listof cell-value?)))]
          [get-sheet-name-cols (-> string? (listof (listof cell-value?)))]
          [get-sheet-*name*-cols (-> string? (listof (listof cell-value?)))]

          [set-cols! (-> (listof (listof cell-value?)) void?)]
          [set-sheet-ref-cols! (-> natural? (listof (listof cell-value?)) void?)]
          [set-sheet-name-cols! (-> string? (listof (listof cell-value?)) void?)]
          [set-sheet-*name*-cols! (-> string? (listof (listof cell-value?)) void?)]

          [get-range-values (-> string? (listof cell-value?))]
          [get-sheet-ref-range-values (-> natural? string? (listof cell-value?))]
          [get-sheet-name-range-values (-> string? string? (listof cell-value?))]
          [get-sheet-*name*-range-values (-> string? string? (listof cell-value?))]

          [squash-shared-strings-map (-> void?)]
          ))

(define (get-sheet-dimension)
  (DATA-SHEET-dimension (*CURRENT_SHEET*)))

(define (get-rows-count)
  (car (range->capacity (DATA-SHEET-dimension (*CURRENT_SHEET*)))))
(define (get-sheet-ref-rows-count sheet_index) (with-sheet-ref sheet_index (lambda ()  (get-rows-count))))
(define (get-sheet-name-rows-count sheet_name) (with-sheet-name sheet_name (lambda () (get-rows-count))))
(define (get-sheet-*name*-rows-count search_sheet_name) (with-sheet-*name* search_sheet_name (lambda () (get-rows-count))))

(define (get-cols-count)
  (cdr (range->capacity (DATA-SHEET-dimension (*CURRENT_SHEET*)))))
(define (get-sheet-ref-cols-count sheet_index) (with-sheet-ref sheet_index (lambda ()  (get-cols-count))))
(define (get-sheet-name-cols-count sheet_name) (with-sheet-name sheet_name (lambda () (get-cols-count))))
(define (get-sheet-*name*-cols-count search_sheet_name) (with-sheet-*name* search_sheet_name (lambda () (get-cols-count))))

(define (get-row-cells row_index)
  (let* ([range_row_col (range->row_col_pair (DATA-SHEET-dimension (*CURRENT_SHEET*)))]
         [start_col (cdar range_row_col)]
         [end_col (cddr range_row_col)])
  (let loop ([loop_col_index start_col]
             [cells '()])
    (if (<= loop_col_index end_col)
        (loop
         (add1 loop_col_index)
         (cons
          (row_col->cell row_index loop_col_index)
          cells))
        (reverse cells)))))
(define (get-sheet-ref-row-cells sheet_index row_index) (with-sheet-ref sheet_index (lambda ()  (get-row-cells row_index))))
(define (get-sheet-name-row-cells sheet_name row_index) (with-sheet-name sheet_name (lambda () (get-row-cells row_index))))
(define (get-sheet-*name*-row-cells search_sheet_name row_index) (with-sheet-*name* search_sheet_name (lambda () (get-row-cells row_index))))

(define (get-col-cells col)
  (cond
   [(string? col)
    (get-col-number-cells (col_abc->number col))]
   [(natural? col)
    (get-col-number-cells col)]))

(define (get-col-number-cells col_index)
  (let* ([range_row_col (range->row_col_pair (DATA-SHEET-dimension (*CURRENT_SHEET*)))]
         [start_row (caar range_row_col)]
         [end_row (cadr range_row_col)])

  (let loop ([loop_row_index start_row]
             [cells '()])
    (if (<= loop_row_index end_row)
        (loop
         (add1 loop_row_index)
         (cons
          (row_col->cell loop_row_index col_index)
          cells))
        (reverse cells)))))
(define (get-sheet-ref-col-cells sheet_index col_index) (with-sheet-ref sheet_index (lambda ()  (get-col-cells col_index))))
(define (get-sheet-name-col-cells sheet_name col_index) (with-sheet-name sheet_name (lambda () (get-col-cells col_index))))
(define (get-sheet-*name*-col-cells search_sheet_name col_index) (with-sheet-*name* search_sheet_name (lambda () (get-col-cells col_index))))

(define (get-cell cell)
  (hash-ref (DATA-SHEET-cell->value_hash (*CURRENT_SHEET*)) cell ""))
(define (get-sheet-ref-cell sheet_index cell) (with-sheet-ref sheet_index (lambda ()  (get-cell cell))))
(define (get-sheet-name-cell sheet_name cell) (with-sheet-name sheet_name (lambda () (get-cell cell))))
(define (get-sheet-*name*-cell search_sheet_name cell) (with-sheet-*name* search_sheet_name (lambda () (get-cell cell))))

(define (set-cell! cell value)
  (hash-set! (DATA-SHEET-cell->value_hash (*CURRENT_SHEET*)) cell value))
(define (set-sheet-ref-cell! sheet_index cell value) (with-sheet-ref sheet_index (lambda ()  (set-cell! cell value))))
(define (set-sheet-name-cell! sheet_name cell value) (with-sheet-name sheet_name (lambda () (set-cell! cell value))))
(define (set-sheet-*name*-cell! search_sheet_name cell value) (with-sheet-*name* search_sheet_name (lambda () (set-cell! cell value))))

(define (get-row row_index)
  (map
   (lambda (cell)
     (get-cell cell))
   (get-row-cells row_index)))
(define (get-sheet-ref-row sheet_index row_index) (with-sheet-ref sheet_index (lambda ()  (get-row row_index))))
(define (get-sheet-name-row sheet_name row_index) (with-sheet-name sheet_name (lambda () (get-row row_index))))
(define (get-sheet-*name*-row search_sheet_name row_index) (with-sheet-*name* search_sheet_name (lambda () (get-row row_index))))

(define (set-row! row_index cell_list)
  (let loop ([cell_strs (get-row-cells row_index)]
             [cell_values cell_list])
    (when (not (null? cell_strs))
          (set-cell! (car cell_strs) (if (null? cell_values) "" (car cell_values)))
          (loop (cdr cell_strs) (cdr cell_values)))))
(define (set-sheet-ref-row! sheet_index row_index cell_list) (with-sheet-ref sheet_index (lambda ()  (set-row! row_index cell_list))))
(define (set-sheet-name-row! sheet_name row_index cell_list) (with-sheet-name sheet_name (lambda () (set-row! row_index cell_list))))
(define (set-sheet-*name*-row! search_sheet_name row_index cell_list) (with-sheet-*name* search_sheet_name (lambda () (set-row! row_index cell_list))))

(define (get-rows)
  (let* ([range_row_col (range->row_col_pair (DATA-SHEET-dimension (*CURRENT_SHEET*)))]
         [start_row (caar range_row_col)]
         [end_row (cadr range_row_col)])
    (let loop ([loop_row_index start_row]
               [rows '()])
    (if (<= loop_row_index end_row)
        (loop
         (add1 loop_row_index)
         (cons
          (get-row loop_row_index)
          rows))
        (reverse rows)))))
(define (get-sheet-ref-rows sheet_index) (with-sheet-ref sheet_index (lambda ()  (get-rows))))
(define (get-sheet-name-rows sheet_name) (with-sheet-name sheet_name (lambda () (get-rows))))
(define (get-sheet-*name*-rows search_sheet_name) (with-sheet-*name* search_sheet_name (lambda () (get-rows))))

(define (set-rows! rows)
  (let* ([range_row_col (range->row_col_pair (DATA-SHEET-dimension (*CURRENT_SHEET*)))]
         [start_row (caar range_row_col)]
         [end_row (cadr range_row_col)])
    (let loop ([loop_row_index 0]
               [actual_loop_row_index start_row])
      (when (<= actual_loop_row_index end_row)
        (set-row! actual_loop_row_index (list-ref rows loop_row_index))
        (loop (add1 loop_row_index) (add1 actual_loop_row_index))))))
(define (set-sheet-ref-rows! sheet_index rows) (with-sheet-ref sheet_index (lambda ()  (set-rows! rows))))
(define (set-sheet-name-rows! sheet_name rows) (with-sheet-name sheet_name (lambda () (set-rows! rows))))
(define (set-sheet-*name*-rows! search_sheet_name rows) (with-sheet-*name* search_sheet_name (lambda () (set-rows! rows))))

(define (get-col col_index)
  (map
   (lambda (cell)
     (get-cell cell))
   (get-col-cells col_index)))
(define (get-sheet-ref-col sheet_index col_index) (with-sheet-ref sheet_index (lambda ()  (get-col col_index))))
(define (get-sheet-name-col sheet_name col_index) (with-sheet-name sheet_name (lambda () (get-col col_index))))
(define (get-sheet-*name*-col search_sheet_name col_index) (with-sheet-*name* search_sheet_name (lambda () (get-col col_index))))

(define (set-col! col_index cell_list)
  (let loop ([cell_strs (get-col-cells col_index)]
             [cell_values cell_list])
    (when (not (null? cell_strs))
          (set-cell! (car cell_strs) (if (null? cell_values) "" (car cell_values)))
          (loop (cdr cell_strs) (cdr cell_values)))))
(define (set-sheet-ref-col! sheet_index col_index cell_list) (with-sheet-ref sheet_index (lambda ()  (set-col! col_index cell_list))))
(define (set-sheet-name-col! sheet_name col_index cell_list) (with-sheet-name sheet_name (lambda () (set-col! col_index cell_list))))
(define (set-sheet-*name*-col! search_sheet_name col_index cell_list) (with-sheet-*name* search_sheet_name (lambda () (set-col! col_index cell_list))))

(define (get-cols)
  (let* ([range_row_col (range->row_col_pair (DATA-SHEET-dimension (*CURRENT_SHEET*)))]
         [start_col (cdar range_row_col)]
         [end_col (cddr range_row_col)])
    (let loop ([loop_col_index start_col]
               [cols '()])
      (if (<= loop_col_index end_col)
          (loop
           (add1 loop_col_index)
           (cons
            (get-col loop_col_index)
            cols))
          (reverse cols)))))
(define (get-sheet-ref-cols sheet_index) (with-sheet-ref sheet_index (lambda ()  (get-cols))))
(define (get-sheet-name-cols sheet_name) (with-sheet-name sheet_name (lambda () (get-cols))))
(define (get-sheet-*name*-cols search_sheet_name) (with-sheet-*name* search_sheet_name (lambda () (get-cols))))

(define (set-cols! cols)
  (let* ([range_row_col (range->row_col_pair (DATA-SHEET-dimension (*CURRENT_SHEET*)))]
         [start_col (cdar range_row_col)]
         [end_col (cddr range_row_col)])
    (let loop ([loop_col_index 0]
               [actual_loop_col_index start_col])
      (when (<= actual_loop_col_index end_col)
          (set-col! actual_loop_col_index (list-ref cols loop_col_index))
          (loop (add1 loop_col_index) (add1 actual_loop_col_index))))))
(define (set-sheet-ref-cols! sheet_index cols) (with-sheet-ref sheet_index (lambda ()  (set-cols! cols))))
(define (set-sheet-name-cols! sheet_name cols) (with-sheet-name sheet_name (lambda () (set-cols! cols))))
(define (set-sheet-*name*-cols! search_sheet_name cols) (with-sheet-*name* search_sheet_name (lambda () (set-cols! cols))))

(define (get-range-values range_str)
  (map
   (lambda (cell)
     (get-cell cell))
   (cell_range->cell_list range_str)))
(define (get-sheet-ref-range-values sheet_index range_str) (with-sheet-ref sheet_index (lambda ()  (get-range-values range_str))))
(define (get-sheet-name-range-values sheet_name range_str) (with-sheet-name sheet_name (lambda () (get-range-values range_str))))
(define (get-sheet-*name*-range-values search_sheet_name range_str) (with-sheet-*name* search_sheet_name (lambda () (get-range-values range_str))))

(define (squash-shared-strings-map)
  (let ([shared_string->index_map (XLSX-shared_string->index_map (*XLSX*))]
        [shared_index->string_map (XLSX-shared_index->string_map (*XLSX*))])

    (hash-clear! shared_string->index_map)
    (hash-clear! shared_index->string_map)

    (let loop ([sheets (XLSX-sheet_list (*XLSX*))]
               [sheet_index 0]
               [sheet_string_index 0])

      (when (not (null? sheets))
            (if (DATA-SHEET? (car sheets))
                (loop (cdr sheets) (add1 sheet_index)
                      (with-sheet-ref
                       sheet_index
                       (lambda ()
                         (let loop-row ([rows (get-rows)]
                                        [row_string_index sheet_string_index])
                           (if (not (null? rows))
                               (loop-row
                                (cdr rows)
                                (let loop-cell ([row_cells (car rows)]
                                                [cell_string_index row_string_index])
                                  (if (not (null? row_cells))
                                      (let ([cell_value (car row_cells)])
                                        (if (string? cell_value)
                                            (if (not (hash-has-key? shared_string->index_map cell_value))
                                                (begin
                                                  (hash-set! shared_string->index_map cell_value cell_string_index)
                                                  (hash-set! shared_index->string_map cell_string_index cell_value)
                                                  (loop-cell (cdr row_cells) (add1 cell_string_index)))
                                                (loop-cell (cdr row_cells) cell_string_index))
                                            (loop-cell (cdr row_cells) cell_string_index)))
                                      cell_string_index)))
                               row_string_index)))))
                (loop (cdr sheets) (add1 sheet_index) sheet_string_index))))))
