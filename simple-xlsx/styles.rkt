#lang racket

(provide (contract-out
          [add-cell-style (-> string? (or/c list? (listof hash?)) void?)]
          [add-cell-range-style (-> string? list? void?)]
          [add-cells-style (-> (listof string?) list? void?)]
          [add-row-style (-> string? list? void?)]
          [add-col-style (-> string? list? void?)]
          ))

(require racket/date)

(require "xlsx/xlsx.rkt")
(require "sheet/sheet.rkt")
(require "lib/lib.rkt")
(require "lib/dimension.rkt")

(define (add-row-style row_range style_list)
  (let ([affected_row_range (check-row-range row_range)]
        [new_style_hashes (style_list->style_hash style_list)]
        [sheet_row->style_index_map (DATA-SHEET-row->style_index_map (*CURRENT_SHEET*))]
        [sheet_row->cells_map (DATA-SHEET-row->cells_map (*CURRENT_SHEET*))])
    (let loop-row ([loop_row_index (car affected_row_range)])
      (when (<= loop_row_index (cdr affected_row_range))
            (add-cells-style (set->list (hash-ref sheet_row->cells_map loop_row_index (set))) new_style_hashes)
            (let* ([exist_style_index (hash-ref sheet_row->style_index_map loop_row_index 0)]
                   [new_style_index (add-style new_style_hashes exist_style_index)])
              (hash-set! sheet_row->style_index_map loop_row_index new_style_index))
            (loop-row (add1 loop_row_index))))))

(define (add-col-style col_range style_list)
  (let ([affected_col_range (check-col-range col_range)]
        [new_style_hashes (style_list->style_hash style_list)]
        [sheet_col->style_index_map (DATA-SHEET-col->style_index_map (*CURRENT_SHEET*))]
        [sheet_col->cells_map (DATA-SHEET-col->cells_map (*CURRENT_SHEET*))])
    (let loop-col ([loop_col_index (car affected_col_range)])
      (when (<= loop_col_index (cdr affected_col_range))
            (add-cells-style (set->list (hash-ref sheet_col->cells_map loop_col_index (set))) new_style_hashes)
            (let* ([exist_style_index (hash-ref sheet_col->style_index_map loop_col_index 0)]
                   [new_style_index (add-style new_style_hashes exist_style_index)])
              (hash-set! sheet_col->style_index_map loop_col_index new_style_index))
            (loop-col (add1 loop_col_index))))))

(define (add-cell-style cell new_style_hashes_or_list)
  (let ([new_style_hashes (if ((listof hash?) new_style_hashes_or_list) new_style_hashes_or_list (style_list->style_hash new_style_hashes_or_list))]
        [sheet_cell->style_index_map (DATA-SHEET-cell->style_index_map (*CURRENT_SHEET*))]
        [sheet_row->cells_map (DATA-SHEET-row->cells_map (*CURRENT_SHEET*))]
        [sheet_col->cells_map (DATA-SHEET-col->cells_map (*CURRENT_SHEET*))])

    (let* ([exist_style_index (hash-ref sheet_cell->style_index_map cell 0)]
           [row_col (cell->row_col cell)]
           [row_index (car row_col)]
           [col_index (cdr row_col)])
      (hash-set! sheet_row->cells_map row_index (set-add (hash-ref sheet_row->cells_map row_index (set)) cell))
      (hash-set! sheet_col->cells_map col_index (set-add (hash-ref sheet_col->cells_map col_index (set)) cell))
      
      (let ([new_style_index (add-style new_style_hashes exist_style_index)])
        (hash-set! sheet_cell->style_index_map cell new_style_index)))))

(define (add-cells-style cells new_style_hashes)
  (let loop-cell ([loop_cells cells])
    (when (not (null? loop_cells))
      (add-cell-style (car loop_cells) new_style_hashes)
      (loop-cell (cdr loop_cells)))))

(define (add-cell-range-style cell_range style_list)
  (when (not (null? cell_range))
        (let ([affected_cell_range (check-cell-range cell_range)]
              [new_style_hashes (style_list->style_hash style_list)]
              [sheet_cell->style_index_map (DATA-SHEET-cell->style_index_map (*CURRENT_SHEET*))]
              [sheet_row->cells_map (DATA-SHEET-row->cells_map (*CURRENT_SHEET*))]
              [sheet_col->cells_map (DATA-SHEET-col->cells_map (*CURRENT_SHEET*))])

            (let* ([range_items (regexp-match #rx"^([A-Z]+)([0-9]+)-([A-Z]+)([0-9]+)$" affected_cell_range)]
                   [start_col_index (col_abc->number (list-ref range_items 1))]
                   [start_row_index (string->number (list-ref range_items 2))]
                   [end_col_index (col_abc->number (list-ref range_items 3))]
                   [end_row_index (string->number (list-ref range_items 4))])

              (add-cells-style
               (let range-loop ([loop_row_index start_row_index]
                                [loop_col_index start_col_index]
                                [cells '()])
                 (if (<= loop_row_index end_row_index)
                     (if (<= loop_col_index end_col_index)
                         (range-loop loop_row_index (add1 loop_col_index) (cons (row_col->cell loop_row_index loop_col_index) cells))
                         (range-loop (add1 loop_row_index) start_col_index cells))
                     (reverse cells)))
               new_style_hashes)))))

(define (style_list->style_hash style_list)
  (let ([style_hash (make-hash)]
        [font_style_hash (make-hash)]
        [num_style_hash (make-hash)]
        [fill_style_hash (make-hash)]
        [border_style_hash (make-hash)]
        [alignment_style_hash (make-hash)])

    (let loop ([styles style_list])
      (when (not (null? styles))
            (let ([style_pair (car styles)])
              (when (and (pair? style_pair) (symbol? (car style_pair)))
                    (cond
                     [(eq? (car style_pair) 'backgroundColor)
                      (hash-set! fill_style_hash 'fgColor (cdr style_pair))]
                     [(or
                       (eq? (car style_pair) 'fontSize)
                       (eq? (car style_pair) 'fontColor)
                       (eq? (car style_pair) 'fontName))
                      (hash-set! font_style_hash (car style_pair) (cdr style_pair))]
                     [(or
                       (eq? (car style_pair) 'numberPrecision)
                       (eq? (car style_pair) 'numberPercent)
                       (eq? (car style_pair) 'numberThousands)
                       (eq? (car style_pair) 'dateFormat)
                       (eq? (car style_pair) 'formatCode))
                      (hash-set! num_style_hash (car style_pair) (cdr style_pair))]
                     [(or
                       (eq? (car style_pair) 'borderDirection)
                       (eq? (car style_pair) 'borderStyle)
                       (eq? (car style_pair) 'borderColor))
                      (hash-set! border_style_hash (car style_pair) (cdr style_pair))]
                     [(or
                       (eq? (car style_pair) 'horizontalAlign)
                       (eq? (car style_pair) 'verticalAlign))
                      (hash-set! alignment_style_hash (car style_pair) (cdr style_pair))])
                    
                    (hash-set! style_hash (car style_pair) (cdr style_pair))
                    (loop (cdr styles))))))
    
    (list style_hash font_style_hash num_style_hash fill_style_hash border_style_hash alignment_style_hash)))

(define (add-style new_style_hashes exist_style_index)
  (let ([style_hash (list-ref new_style_hashes 0)]
        [font_style_hash (list-ref new_style_hashes 1)]
        [num_style_hash (list-ref new_style_hashes 2)]
        [fill_style_hash (list-ref new_style_hashes 3)]
        [border_style_hash (list-ref new_style_hashes 4)]
        [alignment_style_hash (list-ref new_style_hashes 5)])

    (let ([xlsx_style_hash->index_map (XLSX-style_hash->index_map (*CURRENT_XLSX*))]
          [xlsx_style_index->hash_map (XLSX-style_index->hash_map (*CURRENT_XLSX*))]
          [font_style_hash->index_map (XLSX-font_style_hash->index_map (*CURRENT_XLSX*))]
          [font_style_index->hash_map (XLSX-font_style_index->hash_map (*CURRENT_XLSX*))]
          [num_style_hash->index_map (XLSX-num_style_hash->index_map (*CURRENT_XLSX*))]
          [num_style_index->hash_map (XLSX-num_style_index->hash_map (*CURRENT_XLSX*))]
          [fill_style_hash->index_map (XLSX-fill_style_hash->index_map (*CURRENT_XLSX*))]
          [fill_style_index->hash_map (XLSX-fill_style_index->hash_map (*CURRENT_XLSX*))]
          [border_style_hash->index_map (XLSX-border_style_hash->index_map (*CURRENT_XLSX*))]
          [border_style_index->hash_map (XLSX-border_style_index->hash_map (*CURRENT_XLSX*))]
          [alignment_style_hash->index_map (XLSX-alignment_style_hash->index_map (*CURRENT_XLSX*))]
          [alignment_style_index->hash_map (XLSX-alignment_style_index->hash_map (*CURRENT_XLSX*))])

      (let* ([combined_style_hash (hash-copy (hash-ref xlsx_style_index->hash_map exist_style_index (make-hash)))])
        (hash-for-each
         style_hash
         (lambda (k v)
           (hash-set! combined_style_hash k v)))
        
        (if (hash-has-key? xlsx_style_hash->index_map combined_style_hash)
            (hash-ref xlsx_style_hash->index_map combined_style_hash)
            (let ([new_fill_style_hash (make-hash)]
                  [new_font_style_hash (make-hash)]
                  [new_num_style_hash (make-hash)]
                  [new_border_style_hash (make-hash)]
                  [new_alignment_style_hash (make-hash)])
              (hash-for-each
               combined_style_hash
               (lambda (style_k style_v)
                 (cond
                  [(eq? style_k 'backgroundColor)
                   (hash-set! new_fill_style_hash 'fgColor style_v)]
                  [(or
                    (eq? style_k 'fontSize)
                    (eq? style_k 'fontColor)
                    (eq? style_k 'fontName))
                   (hash-set! new_font_style_hash style_k style_v)]
                  [(or
                    (eq? style_k 'numberPrecision)
                    (eq? style_k 'numberPercent)
                    (eq? style_k 'numberThousands)
                    (eq? style_k 'dateFormat)
                    (eq? style_k 'formatCode))
                   (hash-set! new_num_style_hash style_k style_v)]
                  [(or
                    (eq? style_k 'borderDirection)
                    (eq? style_k 'borderStyle)
                    (eq? style_k 'borderColor))
                   (hash-set! new_border_style_hash style_k style_v)]
                  [(or
                    (eq? style_k 'horizontalAlign)
                    (eq? style_k 'verticalAlign))
                   (hash-set! new_alignment_style_hash style_k style_v)])))
                        
              (let ([new_index (add1 (hash-count xlsx_style_hash->index_map))])
                (hash-set! xlsx_style_hash->index_map combined_style_hash new_index)
                (hash-set! xlsx_style_index->hash_map new_index combined_style_hash)
                
                (when (and (> (hash-count new_font_style_hash) 0) (not (hash-has-key? font_style_hash->index_map new_font_style_hash)))
                      (let ([index (add1 (hash-count font_style_hash->index_map))])
                        (hash-set! font_style_hash->index_map new_font_style_hash index)
                        (hash-set! font_style_index->hash_map index new_font_style_hash)))

                (when (and (> (hash-count new_num_style_hash) 0) (not (hash-has-key? num_style_hash->index_map new_num_style_hash)))
                      (let ([index (add1 (hash-count num_style_hash->index_map))])
                        (hash-set! num_style_hash->index_map new_num_style_hash index)
                        (hash-set! num_style_index->hash_map index new_num_style_hash)))

                (when (and (> (hash-count new_fill_style_hash) 0) (not (hash-has-key? fill_style_hash->index_map new_fill_style_hash)))
                      (let ([index (add1 (hash-count fill_style_hash->index_map))])
                        (hash-set! fill_style_hash->index_map new_fill_style_hash index)
                        (hash-set! fill_style_index->hash_map index new_fill_style_hash)))
                
                (when (and (> (hash-count new_border_style_hash) 0) (not (hash-has-key? border_style_hash->index_map new_border_style_hash)))
                      (let ([index (add1 (hash-count border_style_hash->index_map))])
                        (hash-set! border_style_hash->index_map new_border_style_hash index)
                        (hash-set! border_style_index->hash_map index new_border_style_hash)))

                (when (and (> (hash-count new_alignment_style_hash) 0) (not (hash-has-key? alignment_style_hash->index_map new_alignment_style_hash)))
                      (let ([index (add1 (hash-count alignment_style_hash->index_map))])
                        (hash-set! alignment_style_hash->index_map new_alignment_style_hash index)
                        (hash-set! alignment_style_index->hash_map index new_alignment_style_hash)))
                new_index)))))))
