#lang racket

(provide (contract-out
          [write-xlsx-file (-> path-string? void?)]
          [add-data-sheet (-> string? (listof list?) void?)]
          [add-chart-sheet (-> string? (or/c 'LINE 'LINE3D 'BAR 'BAR3D 'PIE 'PIE3D) string? void?)]
          [add-cell-style (-> string? list? void?)]
          ))

(require racket/date)

(require "xlsx/xlsx.rkt")
(require "sheet/sheet.rkt")
(require "lib/lib.rkt")
(require "lib/dimension.rkt")
(require "new/content-type.rkt")
(require "new/_rels/rels.rkt")
(require "new/docProps/docprops-app.rkt")
(require "new/docProps/docprops-core.rkt")
(require "new/xl/_rels/workbook-xml-rels.rkt")
(require "new/xl/printerSettings/printerSettings.rkt")
(require "new/xl/theme/theme.rkt")
(require "new/xl/sharedStrings.rkt")
(require "new/xl/styles/styles.rkt")

(define (write-xlsx-file xlsx_file_name)
  (when (file-exists? xlsx_file_name)
        (delete-file xlsx_file_name))

  (dynamic-wind
      (lambda () (set-XLSX-xlsx_dir! (*CURRENT_XLSX*) (make-temporary-file "xlsx_tmp_~a" 'directory ".")))
      (lambda ()
        ;; [Content_Types].xml
        (write-content-type)

        ;; _rels
        (write-rels)

        ;; docProps-app
        (write-docprops-app)

        ;; docProps-core
        (write-docprops-core)

        ;; xl rels
        (write-workbook-rels)

        ;; printerSettings
        (write-printer-settings)
        
        ;; themme
        (write-theme)

        ;; shared_string
        (write-shared-strings)

        (zip-xlsx xlsx_file_name (XLSX-xlsx_dir (*CURRENT_XLSX*))))
      (lambda ()
        (delete-directory/files (XLSX-xlsx_dir (*CURRENT_XLSX*))))))

(define (add-data-sheet sheet_name sheet_data)
  (check-data-integrity sheet_data)

  (if (not (hash-has-key? (XLSX-sheet_name_index_map (*CURRENT_XLSX*)) sheet_name))
      (let ([seq (add1 (length (XLSX-sheet_list (*CURRENT_XLSX*))))]
            [type_seq (add1 (length (filter (lambda (rec) (DATA-SHEET? rec)) (XLSX-sheet_list (*CURRENT_XLSX*)))))]
            [shared_string_index (hash-count (XLSX-shared_strings_map (*CURRENT_XLSX*)))]
            [rvtsf_map (make-hash)])
        (let row-loop ([rows sheet_data]
                       [row_index 1])
          (when (not (null? rows))
                (let col-loop ([cols (car rows)]
                               [col_index 1])
                  (when (not (null? cols))
                        (cond
                         [(string? (car cols))
                          (set! shared_string_index (add-shared-strings-map (XLSX-shared_strings_map (*CURRENT_XLSX*)) (car cols) shared_string_index))
                          (hash-set! rvtsf_map
                                     (row_col->dimension row_index col_index) ;; r
                                     (list
                                      shared_string_index ;; v t s f
                                      "s"
                                      #f
                                      #f))
                          shared_string_index]
                         [(date? (car cols))
                          (let ([date_v (date->oa_date_number (car cols))])
                            (hash-set! rvtsf_map
                                       (row_col->dimension row_index col_index) ;; r
                                       (list
                                        date_v ;; v t s f
                                        "n"
                                        #f
                                        #f))
                            date_v)]
                         [(number? (car cols))
                          (hash-set! rvtsf_map
                                     (row_col->dimension row_index col_index) ;; r
                                     (list
                                      (car cols) ;; v t s f
                                      "n"
                                      #f
                                      #f))
                          (car cols)])

                      (col-loop (cdr cols) (add1 col_index))))

                (row-loop (cdr rows) (add1 row_index))))

        (let* ([sheet_index (XLSX-sheet_count (*CURRENT_XLSX*))]
               [id (number->string (add1 sheet_index))]
               [rId (format "rId~a" id)]
               [rel (format "worksheets/sheet~a" id)])

          (set-XLSX-sheet_count! (*CURRENT_XLSX*) (add1 (XLSX-sheet_count (*CURRENT_XLSX*))))
          
          (set-XLSX-sheet_list! (*CURRENT_XLSX*)
                                `(,@(XLSX-sheet_list (*CURRENT_XLSX*))
                                  ,(new-data-sheet
                                    (get-dimension sheet_data)
                                    rvtsf_map)))

          (hash-set! (XLSX-sheet_index_id_map (*CURRENT_XLSX*)) sheet_index id)
          (hash-set! (XLSX-sheet_index_name_map (*CURRENT_XLSX*)) sheet_index sheet_name)
          (hash-set! (XLSX-sheet_name_index_map (*CURRENT_XLSX*)) sheet_name sheet_index)
          (hash-set! (XLSX-sheet_index_rid_map (*CURRENT_XLSX*)) sheet_index rId)
          (hash-set! (XLSX-sheet_rid_rel_map (*CURRENT_XLSX*)) rId rel)
          (hash-set! (XLSX-sheet_index_rel_map (*CURRENT_XLSX*)) sheet_index rel)))
      (error (format "duplicate sheet name[~a]" sheet_name))))

(define (add-chart-sheet sheet_name chart_type topic)
  (if (not (hash-has-key? (XLSX-sheet_name_index_map (*CURRENT_XLSX*)) sheet_name))
        (let* ([sheet_index (XLSX-sheet_count (*CURRENT_XLSX*))]
               [id (number->string (add1 sheet_index))]
               [rId (format "rId~a" id)]
               [rel (format "chartsheets/sheet~a" id)])

          (set-XLSX-sheet_count! (*CURRENT_XLSX*) (add1 sheet_index))
          (set-XLSX-sheet_list! (*CURRENT_XLSX*) `(,@(XLSX-sheet_list (*CURRENT_XLSX*)) ,(new-chart-sheet chart_type topic)))
          (hash-set! (XLSX-sheet_index_id_map (*CURRENT_XLSX*)) sheet_index id)
          (hash-set! (XLSX-sheet_index_name_map (*CURRENT_XLSX*)) sheet_index sheet_name)
          (hash-set! (XLSX-sheet_name_index_map (*CURRENT_XLSX*)) sheet_name sheet_index)
          (hash-set! (XLSX-sheet_index_rid_map (*CURRENT_XLSX*)) sheet_index rId)
          (hash-set! (XLSX-sheet_rid_rel_map (*CURRENT_XLSX*)) rId rel)
          (hash-set! (XLSX-sheet_index_rel_map (*CURRENT_XLSX*)) sheet_index rel))
      (error (format "duplicate sheet name[~a]" sheet_name))))

(define (add-cell-style cell_range style_list)
  (let ([formated_cell_range (check-cell-range cell_range)]
        [style_hash (make-hash)]
        [font_style_hash (make-hash)]
        [num_style_hash (make-hash)]
        [fill_style_hash (make-hash)]
        [border_style_hash (make-hash)]
        [alignment_style_hash (make-hash)]
        [style_group_hash (make-hash)]
        )

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
                    (hash-set! style_hash (caar styles) (cdar styles)))
              (loop (cdr styles)))))
    
    (hash-set! style_group_hash 'font font_style_hash)
    (hash-set! style_group_hash 'num num_style_hash)
    (hash-set! style_group_hash 'fill fill_style_hash)
    (hash-set! style_group_hash 'border border_style_hash)
    (hash-set! style_group_hash 'alignment alignment_style_hash)
    
    (let ([xlsx_style_hash->index_map (XLSX-style_hash->index_map (*CURRENT_XLSX*))]
          [xlsx_style_index->hash_map (XLSX-style_index->hash_map (*CURRENT_XLSX*))]
          [sheet_cell->style_index_map (DATA-SHEET-cell->style_index_map (*CURRENT_SHEET*))]
          [font_style_hash->index_map (XLSX-font_style_hash->index_map (*CURRENT_XLSX*))]
          [font_style_index->hash_map (XLSX-font_style_index->hash_map (*CURRENT_XLSX*))]
          [num_style_hash->index_map (XLSX-num_style_hash->index_map (*CURRENT_XLSX*))]
          [num_style_index->hash_map (XLSX-num_style_index->hash_map (*CURRENT_XLSX*))]
          [fill_style_hash->index_map (XLSX-fill_style_hash->index_map (*CURRENT_XLSX*))]
          [fill_style_index->hash_map (XLSX-fill_style_index->hash_map (*CURRENT_XLSX*))]
          [border_style_hash->index_map (XLSX-border_style_hash->index_map (*CURRENT_XLSX*))]
          [border_style_index->hash_map (XLSX-border_style_index->hash_map (*CURRENT_XLSX*))]
          [alignment_style_hash->index_map (XLSX-alignment_style_hash->index_map (*CURRENT_XLSX*))]
          [alignment_style_index->hash_map (XLSX-alignment_style_index->hash_map (*CURRENT_XLSX*))]
          )

      (let* ([range_items (regexp-match #rx"^([A-Z]+)([0-9]+)-([A-Z]+)([0-9]+)$" formated_cell_range)]
             [start_col_index (col_abc->number (list-ref range_items 1))]
             [start_row_index (string->number (list-ref range_items 2))]
             [end_col_index (col_abc->number (list-ref range_items 3))]
             [end_row_index (string->number (list-ref range_items 4))])

        (let range-loop ([loop_col_index start_col_index]
                         [loop_row_index start_row_index])

          (when (and (<= loop_col_index end_col_index) (<= loop_row_index end_row_index))

            (let ([cell (row_col->dimension loop_row_index loop_col_index)])
              (let ([new_style_hash
                     (if (hash-has-key? sheet_cell->style_index_map cell)
                       (let ([cell_style_hash (hash-copy (hash-ref xlsx_style_index->hash_map (hash-ref sheet_cell->style_index_map cell)))]
                             [cell_font_style_hash (hash-copy (hash-ref style_group_hash 'font))]
                             [cell_num_style_hash (hash-copy (hash-ref style_group_hash 'num))]
                             [cell_fill_style_hash (hash-copy (hash-ref style_group_hash 'fill))]
                             [cell_border_style_hash (hash-copy (hash-ref style_group_hash 'border))]
                             [cell_alignment_style_hash (hash-copy (hash-ref style_group_hash 'alignment))])
                         (hash-for-each
                          style_hash
                          (lambda (style_k style_v)
                            (cond
                             [(eq? style_k 'backgroundColor)
                              (hash-set! cell_fill_style_hash 'fgColor (cdr style_pair))]
                             [(or
                               (eq? (car style_pair) 'fontSize)
                               (eq? (car style_pair) 'fontColor)
                               (eq? (car style_pair) 'fontName))
                              (hash-set! cell_font_style_hash (car style_pair) (cdr style_pair))]
                             [(or
                               (eq? (car style_pair) 'numberPrecision)
                               (eq? (car style_pair) 'numberPercent)
                               (eq? (car style_pair) 'numberThousands)
                               (eq? (car style_pair) 'dateFormat)
                               (eq? (car style_pair) 'formatCode))
                              (hash-set! cell_num_style_hash (car style_pair) (cdr style_pair))]
                             [(or
                               (eq? (car style_pair) 'borderDirection)
                               (eq? (car style_pair) 'borderStyle)
                               (eq? (car style_pair) 'borderColor))
                              (hash-set! cell_border_style_hash (car style_pair) (cdr style_pair))]
                             [(or
                               (eq? (car style_pair) 'horizontalAlign)
                               (eq? (car style_pair) 'verticalAlign))
                              (hash-set! cell_alignment_style_hash (car style_pair) (cdr style_pair))])
                            (hash-set! style_hash (caar styles) (cdar styles)))

                            (hash-set! cell_style_hash style_k style_v)
                            ))
                         cell_style_hash)])
                
                  (if (hash-has-key? xlsx_style_hash->index_map new_style_hash)
                      (hash-set! sheet_cell->style_index_map cell (hash-ref xlsx_style_hash->index_map new_style_hash))
                      (let ([new_index (add1 (hash-count xlsx_style_hash->index_map))])
                        (hash-set! xlsx_style_hash->index_map new_style_hash new_index)
                        (hash-set! xlsx_style_index->hash_map new_index new_style_hash)
                        (hash-set! sheet_cell->style_index_map cell new_index)))

                  (when (and (> (hash-count cell_font_style_hash) 0) (not (hash-has-key? font_style_hash->index_map cell_font_style_hash)))
                        (let ([index (add1 (hash-count cell_font_style_hash->index_map))])
                          (hash-set! font_style_hash->index_map cell_font_style_hash index)
                          (hash-set! font_style_index->hash_map index cell_font_style_hash)))

                  (when (and (> (hash-count cell_num_style_hash) 0) (not (hash-has-key? cell_num_style_hash->index_map cell_num_style_hash)))
                        (let ([index (add1 (hash-count cell_num_style_hash->index_map))])
                          (hash-set! num_style_hash->index_map cell_num_style_hash index)
                          (hash-set! num_style_index->hash_map index cell_num_style_hash)))

                  (when (and (> (hash-count cell_fill_style_hash) 0) (not (hash-has-key? cell_fill_style_hash->index_map cell_fill_style_hash)))
                        (let ([index (add1 (hash-count cell_fill_style_hash->index_map))])
                          (hash-set! fill_style_hash->index_map cell_fill_style_hash index)
                          (hash-set! fill_style_index->hash_map index cell_fill_style_hash)))

                  (when (and (> (hash-count cell_border_style_hash) 0) (not (hash-has-key? cell_border_style_hash->index_map cell_border_style_hash)))
                        (let ([index (add1 (hash-count cell_border_style_hash->index_map))])
                          (hash-set! border_style_hash->index_map cell_border_style_hash index)
                          (hash-set! border_style_index->hash_map index cell_border_style_hash)))

                  (when (and (> (hash-count cell_alignment_style_hash) 0) (not (hash-has-key? cell_alignment_style_hash->index_map cell_alignment_style_hash)))
                        (let ([index (add1 (hash-count cell_alignment_style_hash->index_map))])
                          (hash-set! alignment_style_hash->index_map cell_alignment_style_hash index)
                          (hash-set! alignment_style_index->hash_map index cell_alignment_style_hash))))))
                  ))
            (cond
             [(< loop_col_index end_col_index)
              (range-loop (add1 loop_col_index) loop_row_index)]
             [(< loop_row_index end_row_index)
              (range-loop start_col_index (add1 loop_row_index))])
            )))
