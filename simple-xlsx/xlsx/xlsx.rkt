#lang racket

(require "../lib/lib.rkt")
(require "../lib/dimension.rkt")
(require "../sheet/sheet.rkt")
(require "../style/style.rkt")
(require "../style/border-style.rkt")
(require "../style/fill-style.rkt")
(require "../style/font-style.rkt")
(require "../style/number-style.rkt")

(provide (contract-out
          [struct XLSX
                  (
                   (xlsx_dir path-string?)
                   (sheet_list (listof (or/c DATA-SHEET? CHART-SHEET?)))
                   (shared_string->index_map (hash/c string? natural?))
                   (shared_index->string_map (hash/c natural? string?))
                   (styles STYLES?)
                   )
                  ]
          [with-xlsx (-> procedure? void?)]
          [with-sheet-ref (-> natural? procedure? any)]
          [with-sheet (-> procedure? any)]
          [with-sheet-name (-> string? procedure? any)]
          [with-sheet-*name* (-> string? procedure? any)]
          [*XLSX* (parameter/c (or/c XLSX? #f))]
          [*CURRENT_SHEET* (parameter/c (or/c DATA-SHEET? CHART-SHEET? #f))]
          [*CURRENT_SHEET_INDEX* (parameter/c natural?)]
          [add-data-sheet (->* (string? (listof list?)) (cell?) void?)]
          [add-chart-sheet (-> string?
                               (or/c 'LINE 'LINE3D 'BAR 'BAR3D 'PIE 'PIE3D 'UNKNOWN)
                               string?
                               (listof (list/c string? string? string? string? string?))
                               void?)]
          [get-sheet-name-list (-> (listof string?))]
          ))

(define *XLSX* (make-parameter #f))
(define *CURRENT_SHEET* (make-parameter #f))
(define *CURRENT_SHEET_INDEX* (make-parameter 0))

(struct XLSX
        (
         (xlsx_dir #:mutable)
         (sheet_list #:mutable)
         shared_string->index_map
         shared_index->string_map
         styles))

(define (with-xlsx user_proc)
  (parameterize*
   (
    [*XLSX* (new-xlsx)]
    [*STYLES* (XLSX-styles (*XLSX*))]
    [*STYLE->INDEX_MAP* (STYLES-style->index_map (*STYLES*))]
    [*INDEX->STYLE_MAP* (STYLES-index->style_map (*STYLES*))]
    [*BORDER_STYLE->INDEX_MAP* (STYLES-border_style->index_map (*STYLES*))]
    [*BORDER_INDEX->STYLE_MAP* (STYLES-border_index->style_map (*STYLES*))]
    [*FONT_STYLE->INDEX_MAP* (STYLES-font_style->index_map (*STYLES*))]
    [*FONT_INDEX->STYLE_MAP* (STYLES-font_index->style_map (*STYLES*))]
    [*NUMBER_STYLE->INDEX_MAP* (STYLES-number_style->index_map (*STYLES*))]
    [*NUMBER_INDEX->STYLE_MAP* (STYLES-number_index->style_map (*STYLES*))]
    [*FILL_STYLE->INDEX_MAP* (STYLES-fill_style->index_map (*STYLES*))]
    [*FILL_INDEX->STYLE_MAP* (STYLES-fill_index->style_map (*STYLES*))]
    )
   (user_proc)))

(define (with-sheet-ref sheet_index user_proc)
  (when (< sheet_index (length (XLSX-sheet_list (*XLSX*))))
        (let ([sheet (list-ref (XLSX-sheet_list (*XLSX*)) sheet_index)])
          (cond
           [(DATA-SHEET? sheet)
            (parameterize
             (
              [*CURRENT_SHEET* (list-ref (XLSX-sheet_list (*XLSX*)) sheet_index)]
              [*CURRENT_SHEET_INDEX* sheet_index]
              [*CURRENT_SHEET_STYLE* (hash-ref (STYLES-sheet_index->style_map (*STYLES*)) sheet_index)]
              )
             (user_proc))]
           [(CHART-SHEET? sheet)
            (parameterize
             (
              [*CURRENT_SHEET* (list-ref (XLSX-sheet_list (*XLSX*)) sheet_index)]
              [*CURRENT_SHEET_INDEX* sheet_index]
              )
             (user_proc))]
           [else
            (void)]))))

(define (with-sheet user_proc)
  (with-sheet-ref 0 user_proc))

(define (with-sheet-name sheet_name user_proc)
  (let ([sheet_index (index-of (get-sheet-name-list) sheet_name)])
    (if sheet_index
        (with-sheet-ref sheet_index user_proc)
        (error (format "no such sheet name[~a]!" sheet_name)))))

(define (with-sheet-*name* search_sheet_name user_proc)
  (let loop ([sheet_name_list (get-sheet-name-list)])
    (if (not (null? sheet_name_list))
        (if (regexp-match (regexp search_sheet_name) (car sheet_name_list))
            (with-sheet-name (car sheet_name_list) user_proc)
            (loop (cdr sheet_name_list)))
        (error (format "no such sheet name[*~a*]!" search_sheet_name)))))

(define (new-xlsx) (XLSX "" '() (make-hash) (make-hash) (new-styles)))

(define (add-data-sheet sheet_name sheet_data [start_cell? "A1"])
  (check-data-integrity sheet_data)

  (if (not (member sheet_name (get-sheet-name-list)))
      (let ([sheet_index (length (XLSX-sheet_list (*XLSX*)))]
            [cell->value_map (make-hash)])

        (let ([sheet_style (new-sheet-style)]
              [start_cell_row_col (cell->row_col start_cell?)])
          (let row-loop ([rows sheet_data]
                         [row_index (car start_cell_row_col)])
            (when (not (null? rows))
              (let col-loop ([cols (car rows)]
                             [col_index (cdr start_cell_row_col)])
                (when (not (null? cols))
                  (let ([cell (row_col->cell row_index col_index)])
                    (hash-set! cell->value_map cell (car cols))
                    (col-loop (cdr cols) (add1 col_index)))))

                (row-loop (cdr rows) (add1 row_index))))

          (hash-set! (STYLES-sheet_index->style_map (*STYLES*)) sheet_index sheet_style))

        (set-XLSX-sheet_list! (*XLSX*)
                              `(,@(XLSX-sheet_list (*XLSX*))
                                ,(DATA-SHEET
                                  sheet_name
                                  (capacity->range (cons (length sheet_data) (length (car sheet_data))) start_cell?)
                                  cell->value_map))))
      (error (format "duplicate sheet name[~a]" sheet_name))))

(define (add-chart-sheet sheet_name chart_type topic serial)
  (if (not (member sheet_name (get-sheet-name-list)))
      (set-XLSX-sheet_list! (*XLSX*)
                            `(,@(XLSX-sheet_list (*XLSX*))
                              ,(CHART-SHEET sheet_name chart_type topic serial)))
      (error (format "duplicate sheet name[~a]" sheet_name))))

(define (get-sheet-name-list)
  (let loop ([sheet_list (XLSX-sheet_list (*XLSX*))]
             [sheet_index 0]
             [result_list '()])
    (if (not (null? sheet_list))
        (loop
         (cdr sheet_list)
         (add1 sheet_index)
         (cons
          (cond
           [(DATA-SHEET? (car sheet_list))
            (DATA-SHEET-sheet_name (car sheet_list))]
           [(CHART-SHEET? (car sheet_list))
            (CHART-SHEET-sheet_name (car sheet_list))])
          result_list))
        (reverse result_list))))
