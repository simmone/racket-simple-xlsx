#lang racket

(require "../../../lib/lib.rkt")
(require "../../../xlsx.rkt")

(provide (contract-out
          [write-data-sheet (-> string? (is-a?/c xlsx%) string?)]
          [write-data-sheet-file (-> path-string? (is-a?/c xlsx%) void?)]
          ))

(define (write-data-sheet sheet_name xlsx)
  (let ([color_style_map (make-hash)])
    (let loop ([loop_list color_list]
               [index 1])
      (when (not (null? loop_list))
            (hash-set! color_style_map (car loop_list) index)
            (loop (cdr loop_list))))

    (with-output-to-string
      (lambda ()
        ;; only data sheet
        (let loop ([loop_list sheet_list])
          (when (not (null? loop_list))
                (when (eq? (sheet-type (car loop_list)) 'data)
                      (let* ([sheet (car loop_list)]
                             [data_sheet (sheet-content sheet)]
                             [rows (data-sheet-rows sheet)]
                             [col_count (length (car rows))]
                             [width_hash (data-sheet-width_hash data_sheet)]
                             [color_hash (data-sheet-color_hash data_sheet)]
                             [dimension (if (= (length rows) 0) "A1" (string-append "A1:" (get-dimension rows)))]
                             [is_active (if (= (sheet-seq sheet) 1) "tabSelected=\"1\"" "")]
                             [active_cell (if (null? data_sheet) "" "<selection activeCell=\"A1\" sqref=\"A1\"/>")])

                        (printf "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n")
                        (printf "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\"><dimension ref=\"~a\"/><sheetViews><sheetView ~a workbookViewId=\"0\">~a</sheetView></sheetViews><sheetFormatPr defaultRowHeight=\"13.5\"/>" dimension is_active active_cell)

                        (when (> (hash-count width_hash) 0)
                              (printf "<cols>")
                              
                              (hash-for-each
                               width_hash
                               (lambda (col_range width)
                                 (let* ([items (regexp-match* #rx"([A-Z]+)" col_range)]
                                        [start_index (abc->number (first items))]
                                        [end_index (abc->number (second items))])

                                   (printf "<col min=\"~a\" max=\"~a\" width=\"~a\"/>" start_index end_index width))))
                              (printf "</cols>"))
                        
                        (printf "<sheetData>")
                        
                        (let loop-row ([loop_rows data_sheet]
                                       [row_seq 1])
                          (when (not (null? loop_rows))
                                (printf "<row r=\"~a\">" row_seq)
                                (let ([item_list (car loop_rows)])
                                  (when (not (null? item_list))
                                        (let loop-col ([loop_cols item_list]
                                                       [col_seq 1])
                                          (when (not (null? loop_cols))
                                                (let* ([cell (car loop_cols)]
                                                       [dimension (string-append (number->abc col_seq) (number->string row_seq))]
                                                       [color (range-hash-ref color_hash dimension)]
                                                       [color_index
                                                        (if color (string-append "s=\"" (number->string (hash-ref color_style_map color)) "\"") "")])
                                                  (cond
                                                   [(string? cell)
                                                    (printf "<c r=\"~a\" ~a t=\"s\"><v>~a</v></c>" 
                                                            dimension color_style (hash-ref string_index_map cell))]
                                                   [(number? cell)
                                                    (printf "<c r=\"~a\" ~a><v>~a</v></c>" 
                                                            dimension style (number->string (exact->inexact cell)))]
                                                   [else
                                                    (printf "<c r=\"~a\"><v>0</v></c>" dimension)]))
                                                (loop-col (cdr loop_cols) (add1 col_seq))))))
                                (printf "</row>")
                                (loop-row (cdr loop_rows) (add1 row_seq))))))
                (loop (cdr loop_list))))

        (printf "</sheetData><phoneticPr fontId=\"1\" type=\"noConversion\"/><pageMargins left=\"0.7\" right=\"0.7\" top=\"0.75\" bottom=\"0.75\" header=\"0.3\" footer=\"0.3\"/><pageSetup paperSize=\"9\" orientation=\"portrait\" horizontalDpi=\"200\" verticalDpi=\"200\" r:id=\"rId1\"/></worksheet>")))))

(define (write-data-sheet-file dir xlsx)
  (make-directory* dir)

  (let ([loop_list (get-field sheets xlsx)])
    (when (not (null? loop_list))
          (when (eq? (sheet-type (car loop_list)) 'data)
                (with-output-to-file (build-path dir (string-append "sheet" (number->string (sheet-typeSeq (car loop_list))) ".xml"))
                  #:exists 'replace
                  (lambda ()
                    (printf "~a" (write-data-sheet sheet_name xlsx)))))
          (loop (cdr loop_list)))))
  
