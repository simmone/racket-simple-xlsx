#lang racket

(require "../../../lib/lib.rkt")
(require "../../../xlsx.rkt")

(provide (contract-out
          [write-sheet (-> list? string? void?)]
          [write-sheet-file (-> path-string? list? void?)]
          ))

(define (write-sheet sheet_list) 
  (with-output-to-string
    (lambda ()
      ;; only data sheet
      (let loop ([loop_list sheet_list])
        (when (not (null? loop_list))
              (when (eq? (sheet-type (car loop_list)) 'data)
                    (let* ([sheet (car loop_list)]
                           [data_sheet (sheet-content sheet)]
                           [rows (data-sheet-rows sheet)]
                           [width_hash (data-sheet-width_hash sheet)]
                           [dimension (if (null? sheet_data_list) "A1" (string-append "A1:" (get-dimension data_sheet)))]
                           [is_active (if (= (sheet-seq sheet) 1) "tabSelected=\"1\"" "")]
                           [active_cell (if (null? data_sheet) "" "<selection activeCell=\"A1\" sqref=\"A1\"/>")])
                      (printf "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n")
                      (printf "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\"><dimension ref=\"~a\"/><sheetViews><sheetView ~a workbookViewId=\"0\">~a</sheetView></sheetViews><sheetFormatPr defaultRowHeight=\"13.5\"/>" dimension is_active active_cell)

              (printf "<cols>")

              (let loop-col ([loop_cols 
              (when (not (null? loop_cols))
                    (let* ([col (car loop_cols)]
                           [col_index_range (abc->range (car col))]
                           [col_attr (cdr col)]
                           [col_width (col-attr-width col_attr)]
                           [col_color (col-attr-color col_attr)]
                           [col_style_index (if (hash-has-key? color_style_map col_color) (hash-ref color_style_map col_color) #f)])
                      (printf "<col min=\"~a\" max=\"~a\" width=\"~a\" ~a/>"
                              (car col_index_range)
                              (cdr col_index_range)
                              (exact->inexact (cx-round (/ col_width 8) 2))
                              (if col_style_index 
                                  (begin
                                    (hash-set! col_style_map col_index_range col_style_index)
                                    (string-append "style=\"" (number->string col_style_index)"\""))
                                  "")
                              ))
                    (loop-col (cdr loop_cols))))
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
                                                     [style_index (get-range-ref col_style_map col_seq)]
                                                     [style (if style_index (string-append "s=\"" (number->string style_index) "\"") "")])
                                (cond
                                 [(string? cell)
                                  (printf "<c r=\"~a\" ~a t=\"s\"><v>~a</v></c>" dimension style (hash-ref string_index_map cell))]
                                 [(number? cell)
                                  (printf "<c r=\"~a\" ~a><v>~a</v></c>" dimension style (number->string (exact->inexact cell)))]
                                 [else
                                  (printf "<c r=\"~a\"><v>0</v></c>" dimension)]))
                              (loop-col (cdr loop_cols) (add1 col_seq))))))
              (printf "</row>")
              (loop-row (cdr loop_rows) (add1 row_seq)))))))|</sheetData><phoneticPr fontId="1" type="noConversion"/><pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/><pageSetup paperSize="9" orientation="portrait" horizontalDpi="200" verticalDpi="200" r:id="rId1"/></worksheet>
})

(define (write-sheet-file dir sheet_list)
  (with-output-to-file (build-path dir (string-append "sheet" (number->string sheet_index) ".xml"))
    #:exists 'replace
    (lambda ()
      (printf "~a" (write-sheet sheet_list)))))
  
