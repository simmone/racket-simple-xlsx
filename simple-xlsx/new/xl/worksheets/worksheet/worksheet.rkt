#lang at-exp racket/base

(require racket/port)
(require racket/file)
(require racket/class)
(require racket/list)
(require racket/contract)

(require "../../../../lib/lib.rkt")
(require "../../../../xlsx/xlsx.rkt")
(require "../../../../xlsx/range-lib.rkt")
(require "../../../../xlsx/sheet.rkt")

(provide (contract-out
          [write-data-sheet (-> string? (is-a?/c xlsx%) output-port? string?)]
          [write-data-sheet-file (-> path-string? (is-a?/c xlsx%) void?)]
          [get-col-width-map (-> (listof list?) hash?)]
          ))

(define S string-append)

(define (write-header) @S{
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>

<worksheet
  xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
})

(define (write-dimension dimension) @S{
<dimension ref="@|dimension|"/>
})

(define (write-sheet-views is_active active_cell freeze_range) @S{
<sheetViews>
  <sheetView 
    @|(with-output-to-string
        (lambda ()
          (let ([freeze_rows (car freeze_range)]
                [freeze_cols (cdr freeze_range)])
            (when (and
                  (= freeze_rows 0)
                  (= freeze_cols 0))
              (printf "~a" is_active)))))| workbookViewId="0">@|(with-output-to-string
        (lambda ()
          (let ([freeze_rows (car freeze_range)]
                [freeze_cols (cdr freeze_range)])
            (when (or
                     (> freeze_rows 0)
                     (> freeze_cols 0))

              (printf "<pane")

              (when (> freeze_rows 0)
                (printf " ySplit=\"~a\"" freeze_rows))

              (when (> freeze_cols 0)
                (printf " xSplit=\"~a\"" freeze_cols))

              (printf " topLeftCell=\"~a~a\"" (number->abc (add1 freeze_cols)) (add1 freeze_rows))

              (cond
                [(and (> freeze_rows 0) (= freeze_cols 0))
                 (printf " activePane=\"bottomLeft\" state=\"frozen\" />\n")
                 (printf "    <selection pane=\"bottomLeft\" />\n")]
                [(and (= freeze_rows 0) (> freeze_cols 0))
                 (printf " activePane=\"topRight\" state=\"frozen\" />\n")
                 (printf "    <selection pane=\"topRight\" />\n")]
                [(and (> freeze_rows 0) (> freeze_cols 0))
                 (printf " activePane=\"bottomRight\" state=\"frozen\" />\n")
                 (printf "    <selection pane=\"bottomLeft\" />\n")
                 (printf "    <selection pane=\"topRight\" />\n")
                 (printf "    <selection pane=\"bottomRight\" />\n")])))
    (printf "\n")))|
  </sheetView>
</sheetViews>
})

(define (write-sheet-formatPr) @S{
<sheetFormatPr defaultRowHeight="13.5"/>
})

(define (write-cols-style col_style_list) @S{
<cols>
@|(with-output-to-string
    (lambda ()
      (let loop ([col_styles col_style_list])
        (when (not (null? col_styles))
          (let* ([col_range (caar col_styles)]
                 [val_list (cdar col_styles)]
                 [width (first val_list)]
                 [style (second val_list)]
                 [width_str (if width (format " width=\"~a\"" width) "")]
                 [style_str (if style (format " style=\"~a\"" style) "")])
            (printf "  <col min=\"~a\" max=\"~a\"~a~a/>\n" (car col_range) (cdr col_range) width_str style_str))
          (loop (cdr col_styles))))))|</cols>
})

(define (write-footer) @S{
<phoneticPr fontId="1" type="noConversion"/>

<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>

<pageSetup paperSize="9" orientation="portrait" horizontalDpi="200" verticalDpi="200" r:id="rId1"/>
})

(define (output-sheet dimension is_active active_cell freeze_range cols_str rows_str) @S{
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>

<worksheet
  xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">

@|(prefix-each-line (write-dimension dimension) "  ")|

@|(prefix-each-line (write-sheet-views is_active active_cell freeze_range) "  ")|

@|(prefix-each-line (write-sheet-formatPr) "  ")|

@|(prefix-each-line cols_str "  ")|

  <sheetData>
@|(prefix-each-line rows_str "    ")|  </sheetData>

@|(prefix-each-line (write-footer) "  ")|

</worksheet>
})

(define (get-rows sheet_name xlsx)
  (let* ([string_index_map (send xlsx get-string-index-map)]
         [cell_to_style_index_hash (send xlsx get-cell-to-style-index-map sheet_name)]
         [row_to_style_index_hash (send xlsx get-row-to-style-index-map sheet_name)]
         [col_to_style_index_hash (send xlsx get-col-to-style-index-map sheet_name)]
         [sheet (send xlsx get-sheet-by-name sheet_name)]
         [data_sheet (sheet-content sheet)]
         [height_hash (data-sheet-height_hash data_sheet)]
         [rows (data-sheet-rows data_sheet)]
         [col_count (length (car rows))]
         [span_str (string-append "1:" (number->string col_count))]
         [row_height_map (make-hash)])

    (hash-for-each
     height_hash
     (lambda (row_range height)
       (let ([row_map (range-to-row-hash row_range height)])
         (hash-for-each
          row_map
          (lambda (each_row each_height)
            (hash-set! row_height_map each_row each_height))))))

    (with-output-to-string
      (lambda ()
        (let loop-row ([loop_rows rows]
                       [row_seq 1])
          (when (not (null? loop_rows))
            (let ([style_str ""]
                  [height_str ""])

              (when (hash-has-key? row_to_style_index_hash row_seq)
                (set! style_str (format " s=\"~a\" customFormat=\"1\"" (hash-ref row_to_style_index_hash row_seq))))

              (when (hash-has-key? row_height_map row_seq)
                (set! height_str (format " ht=\"~a\" customHeight=\"1\"" (hash-ref row_height_map row_seq))))

              (printf "<row r=\"~a\" spans=\"~a\"~a~a>\n" row_seq span_str style_str height_str)

              (let ([item_list (car loop_rows)])
                (when (not (null? item_list))
                  (let loop-col ([loop_cols item_list]
                                 [col_seq 1])
                    (when (not (null? loop_cols))
                      (let* ([cell (car loop_cols)]
                             [dimension (string-append (number->abc col_seq) (number->string row_seq))]
                             [style
                                 (cond
                                  [(hash-has-key? cell_to_style_index_hash dimension)
                                   (format " s=\"~a\"" (hash-ref cell_to_style_index_hash dimension))]
                                  [(hash-has-key? row_to_style_index_hash row_seq)
                                   (format " s=\"~a\"" (hash-ref row_to_style_index_hash row_seq))]
                                  [(hash-has-key? col_to_style_index_hash col_seq)
                                   (format " s=\"~a\"" (hash-ref col_to_style_index_hash col_seq))]
                                  [else
                                   ""])])
                        (cond
                         [(string? cell)
                          (printf "  <c r=\"~a\"~a t=\"s\"><v>~a</v></c>\n" 
                                  dimension style (hash-ref string_index_map cell))]
                         [(exact-integer? cell)
                          (printf "  <c r=\"~a\"~a><v>~a</v></c>\n" 
                                  dimension style (number->string (inexact->exact cell)))]
                         [(number? cell)
                          (printf "  <c r=\"~a\"~a><v>~a</v></c>\n" 
                                  dimension style (number->string (exact->inexact cell)))]
                         [else
                          (printf "  <c r=\"~a\"><v>0</v></c>\n" dimension)]))
                      (loop-col (cdr loop_cols) (add1 col_seq))))))
              (printf "</row>\n")
              (when (> (length loop_rows) 1) (printf "\n")))
            (loop-row (cdr loop_rows) (add1 row_seq))))))))

(define (write-data-sheet sheet_name xlsx debug_port)
  (let* ([sheet (send xlsx get-sheet-by-name sheet_name)]
         [data_sheet (sheet-content sheet)]
         [rows (data-sheet-rows data_sheet)]
         [col_width_map (get-col-width-map rows)]
         [col_to_style_index_hash (send xlsx get-col-to-style-index-map sheet_name)]
         [freeze_range (data-sheet-freeze_range data_sheet)]
         [width_hash (data-sheet-width_hash data_sheet)]
         [dimension (if (= (length rows) 0) "A1" (string-append "A1:" (get-dimension rows)))]
         [is_active (if (= (sheet-seq sheet) 1) "tabSelected=\"1\"" "")]
         [active_cell (if (null? data_sheet) "" "<selection activeCell=\"A1\" sqref=\"A1\"/>")]
         [cols_str #f]
         [rows_str #f])
    
    (set! rows_str (get-rows sheet_name xlsx))
    
    (hash-for-each
     width_hash
     (lambda (col_range width)
       (let ([col_map (range-to-col-hash col_range width)])
         (hash-for-each
          col_map
          (lambda (each_col each_width)
            (hash-set! col_width_map each_col each_width))))))

    (set! cols_str (write-cols-style (combine-cols-hash col_width_map col_to_style_index_hash)))

    (output-sheet dimension is_active active_cell freeze_range cols_str rows_str)))

(define (write-data-sheet-file dir xlsx)
  (make-directory* dir)

  (let loop ([loop_list (get-field sheets xlsx)])
    (when (not (null? loop_list))
          (when (eq? (sheet-type (car loop_list)) 'data)
            (let ([debug_port (current-output-port)])
              (with-output-to-file (build-path dir (string-append "sheet" (number->string (sheet-typeSeq (car loop_list))) ".xml"))
                #:exists 'replace
                (lambda ()
                  (printf "~a" (write-data-sheet (sheet-name (car loop_list)) xlsx debug_port))))))
          (loop (cdr loop_list)))))

(define (get-col-width-map rows)
  (let ([col_width_map (make-hash)])
    (let loop-row ([loop_rows rows])
      (when (not (null? loop_rows))
            (let loop-col ([cols (car loop_rows)]
                           [index 1])
              (when (not (null? cols))
                    (let* ([str (if (number? (car cols)) (number->string (car cols)) (car cols))]
                           [str_len (string-length str)]
                           [bytes_len (bytes-length (string->bytes/utf-8 str))]
                           [result_len (+ 2 (- bytes_len (floor (/ (- bytes_len str_len) 2))))])
                      (when (< (hash-ref col_width_map index 0) result_len)
                            (hash-set! col_width_map index result_len)))
                    (loop-col (cdr cols) (add1 index))))
            (loop-row (cdr loop_rows))))
    col_width_map))
