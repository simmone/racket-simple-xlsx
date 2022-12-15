#lang racket

(require simple-xml)

(require "../../xlsx/xlsx.rkt")
(require "../../sheet/sheet.rkt")
(require "../../style/style.rkt")
(require "../../style/sort-styles.rkt")
(require "../../style/set-styles.rkt")
(require "../../style/lib.rkt")
(require "../../lib/dimension.rkt")
(require "../../lib/lib.rkt")

(provide (contract-out
          [to-work-sheet-head (-> list?)]
          [from-work-sheet-head (-> hash? void?)]
          [to-sheet-view (-> list?)]
          [from-sheet-view (-> hash? void?)]
          [to-sheet-views (-> list?)]
          [to-col-width-style (-> list?)]
          [from-col-width-style (-> hash? void?)]
          [to-rows (-> list?)]
          [from-rows (-> hash? void?)]
          [to-merge-cells (-> list?)]
          [from-merge-cells (-> hash? void?)]
          [to-row (-> positive-integer? positive-integer? positive-integer? list?)]
          [from-row (-> hash? positive-integer? void?)]
          [to-cells (-> positive-integer? list?)]
          [to-cell (-> positive-integer? positive-integer? list?)]
          [from-cells (-> hash? positive-integer? void?)]
          [from-cell (-> hash? positive-integer? positive-integer? void?)]
          [work-sheet-tail (-> list?)]
          [to-work-sheet (-> list?)]
          [from-work-sheet (-> hash? void?)]
          [write-worksheets (->* () (path-string?) void?)]
          [read-worksheets (->* () (path-string?) void?)]
          ))

(define (to-work-sheet)
  (append
   (to-work-sheet-head)
   (list (to-sheet-views))
   '(("sheetFormatPr" ("defaultRowHeight" . "13.5")))
   (list (to-col-width-style))
   (list (to-rows))
   (list (to-merge-cells))
   (work-sheet-tail)))

(define (from-work-sheet xml_hash)
  (hash-clear! (DATA-SHEET-cell->value_hash (*CURRENT_SHEET*)))
  (from-work-sheet-head xml_hash)
  (from-sheet-view xml_hash)
  (from-col-width-style xml_hash)
  (from-merge-cells xml_hash)
  (from-rows xml_hash)
  )

(define (to-work-sheet-head)
  (list
   "worksheet"
   '("xmlns" . "http://schemas.openxmlformats.org/spreadsheetml/2006/main")
   '("xmlns:r" . "http://schemas.openxmlformats.org/officeDocument/2006/relationships")
   (list
    "dimension"
    (cons "ref" (DATA-SHEET-dimension (*CURRENT_SHEET*))))))

(define (from-work-sheet-head xml_hash)
  (if (hash-has-key? xml_hash "worksheet1.dimension1.ref")
      (set-DATA-SHEET-dimension! (*CURRENT_SHEET*) (hash-ref xml_hash "worksheet1.dimension1.ref"))
      (set-DATA-SHEET-dimension! (*CURRENT_SHEET*) "A1")))

(define (to-sheet-views)
  (list
   "sheetViews"
   (to-sheet-view)))

(define (to-sheet-view)
  (let ([freeze_rows (car (SHEET-STYLE-freeze_range (*CURRENT_SHEET_STYLE*)))]
        [freeze_cols (cdr (SHEET-STYLE-freeze_range (*CURRENT_SHEET_STYLE*)))])

    (append
     '("sheetView")

     (if (= (*CURRENT_SHEET_INDEX*) 0)
         '(("tabSelected" . "1"))
         '())

     '(("workbookViewId" . "0"))

     (if (not (equal? (SHEET-STYLE-freeze_range (*CURRENT_SHEET_STYLE*)) '(0 . 0)))
         (append
          (list
           (append
            '("pane")

            (if (> freeze_rows 0)
                (list (cons "ySplit" (number->string freeze_rows)))
                '())

            (if (> freeze_cols 0)
                (list (cons "xSplit" (number->string freeze_cols)))
                '())

            (list (cons "topLeftCell" (format "~a~a" (col_number->abc (add1 freeze_cols)) (add1 freeze_rows))))

            (cond
             [(and (> freeze_rows 0) (= freeze_cols 0))
              (list '("activePane" . "bottomLeft") '("state" . "frozen"))]
             [(and (= freeze_rows 0) (> freeze_cols 0))
              (list '("activePane" . "topRight") '("state" . "frozen"))]
             [(and (> freeze_rows 0) (> freeze_cols 0))
              (list '("activePane" . "bottomRight") '("state" . "frozen"))])))

          (cond
           [(and (> freeze_rows 0) (= freeze_cols 0))
            '(("selection" ("pane" . "bottomLeft")))]
           [(and (= freeze_rows 0) (> freeze_cols 0))
            '(("selection" ("pane" . "topRight")))]
           [(and (> freeze_rows 0) (> freeze_cols 0))
            '(("selection" ("pane" . "bottomLeft"))
              ("selection" ("pane" . "topRight"))
              ("selection" ("pane" . "bottomRight")))]))
         '()))))

(define (from-sheet-view xml_hash)
  (set-SHEET-STYLE-freeze_range!
   (*CURRENT_SHEET_STYLE*)
   (cons
    (string->number (hash-ref xml_hash "worksheet1.sheetViews1.sheetView1.pane1.ySplit" "0"))
    (string->number (hash-ref xml_hash "worksheet1.sheetViews1.sheetView1.pane1.xSplit" "0")))))

(define (to-rows)
  (append
   '("sheetData")
   (let* ([range_row_col (range->row_col_pair (DATA-SHEET-dimension (*CURRENT_SHEET*)))]
          [start_row (caar range_row_col)]
          [start_col (cdar range_row_col)]
          [end_row (cadr range_row_col)]
          [end_col (cddr range_row_col)])
     (let loop ([loop_row_index start_row]
                [result_list '()])
       (if (<= loop_row_index end_row)
           (loop
            (add1 loop_row_index)
            (cons
             (to-row loop_row_index start_col end_col)
             result_list))
           (reverse result_list))))))

(define (to-row row_index col_start_index col_end_index)
  (append
   (list
    "row"
    (cons "r" (number->string row_index))
    (cons "spans" (format "~a:~a" col_start_index col_end_index)))
   (if (hash-has-key? (SHEET-STYLE-row->style_map (*CURRENT_SHEET_STYLE*)) row_index)
       (list
        (cons "s" (number->string
                   (add1
                    (hash-ref (STYLES-style->index_map (*STYLES*))
                              (STYLE-hash_code (hash-ref (SHEET-STYLE-row->style_map (*CURRENT_SHEET_STYLE*)) row_index))))))
        (cons "customFormat" "1"))
       '())
   (if (hash-has-key? (SHEET-STYLE-row->height_map (*CURRENT_SHEET_STYLE*)) row_index)
       (list
        (cons "ht" (number->string (hash-ref (SHEET-STYLE-row->height_map (*CURRENT_SHEET_STYLE*)) row_index)))
        (cons "customHeight" "1"))
       '())
   (to-cells row_index)))

(define (to-cells row_index)
  (let* ([range_row_col (range->row_col_pair (DATA-SHEET-dimension (*CURRENT_SHEET*)))]
         [start_row (caar range_row_col)]
         [start_col (cdar range_row_col)]
         [end_row (cadr range_row_col)]
         [end_col (cddr range_row_col)])

    (let loop ([col_index start_col]
               [cell_list '()])
      (if (<= col_index end_col)
          (loop
           (add1 col_index)
           (cons
            (to-cell row_index col_index)
            cell_list))
          (reverse cell_list)))))

(define (to-cell row_index col_index)
  (let ([cell (row_col->cell row_index col_index)])
    (if (hash-has-key? (DATA-SHEET-cell->value_hash (*CURRENT_SHEET*)) cell)
        (let ([cell_value (hash-ref (DATA-SHEET-cell->value_hash (*CURRENT_SHEET*)) cell)])
          (append
           (list
            "c"
            (cons "r" cell))
           (if (hash-has-key? (SHEET-STYLE-cell->style_map (*CURRENT_SHEET_STYLE*)) cell)
               (list
                (cons "s"
                      (number->string
                       (add1
                        (hash-ref (STYLES-style->index_map (*STYLES*))
                                  (STYLE-hash_code (hash-ref (SHEET-STYLE-cell->style_map (*CURRENT_SHEET_STYLE*)) cell)))))))
               '())
           (list
            (if (string? cell_value)
                (cons "t" "s")
                '()))
           (list
            (cond
             [(string? cell_value)
              (list "v" (format "~a" (hash-ref (XLSX-shared_string->index_map (*XLSX*)) cell_value)))]
             [(number? cell_value)
              (list "v" (format "~a" cell_value))]
             [(date? cell_value)
              (list "v" (format "~a" (date->oa_date_number cell_value)))]))))
        '())))

(define (from-rows xml_hash)
  (let* ([rows_count (hash-ref xml_hash "worksheet1.sheetData1.row's count" 0)])

    (let loop-row ([loop_row_index 1])
      (when (<= loop_row_index rows_count)
            (from-row xml_hash loop_row_index)
            (loop-row (add1 loop_row_index))))))

(define (from-row xml_hash loop_row_index)
  (let* ([prefix (format "worksheet1.sheetData1.row~a" loop_row_index)])
    (when (hash-has-key? xml_hash (format "~a.r" prefix))
      (let ([row_index (string->number (hash-ref xml_hash (format "~a.r" prefix)))]
            [row_style (hash-ref xml_hash (format "~a.s" prefix) #f)]
            [row_height (hash-ref xml_hash (format "~a.ht" prefix) #f)])

        (when row_style
          (hash-set! (SHEET-STYLE-row->style_map (*CURRENT_SHEET_STYLE*))
                     row_index
                     (style-from-hash-code (hash-ref (*INDEX->STYLE_MAP*) (string->number row_style)))))

        (when row_height
          (hash-set! (SHEET-STYLE-row->height_map (*CURRENT_SHEET_STYLE*))
                     row_index
                     (string->number row_height))))
    
      (from-cells xml_hash loop_row_index))))

(define (from-cells xml_hash loop_row_index)
  (let* ([cells_count (hash-ref xml_hash (format "worksheet1.sheetData1.row~a.c's count" loop_row_index) 0)])
    (let loop-cell ([loop_cell_index 1])
      (when (<= loop_cell_index cells_count)
            (from-cell xml_hash loop_row_index loop_cell_index)
            (loop-cell (add1 loop_cell_index))))))

(define (from-cell xml_hash loop_row_index loop_cell_index)
  (let* ([prefix (format "worksheet1.sheetData1.row~a" loop_row_index)]
         [cell_ref (format "~a.c~a.r" prefix loop_cell_index)]
         [cell (hash-ref xml_hash cell_ref #f)]
         [cell_type (hash-ref xml_hash (format "~a.c~a.t" prefix loop_cell_index) "number")]
         [cell_style (hash-ref xml_hash (format "~a.c~a.s" prefix loop_cell_index) #f)]
         [cell_value (hash-ref xml_hash (format "~a.c~a.v1" prefix loop_cell_index) #f)])

    (when (and cell cell_value)
      (when cell_style
        (hash-set! (SHEET-STYLE-cell->style_map (*CURRENT_SHEET_STYLE*))
                   cell
                   (style-from-hash-code (hash-ref (*INDEX->STYLE_MAP*) (string->number cell_style)))))
          
      (cond
       [(string=? cell_type "number")
        (hash-set! (DATA-SHEET-cell->value_hash (*CURRENT_SHEET*)) cell (string->number cell_value))]
       [(string=? cell_type "s")
        (hash-set! (DATA-SHEET-cell->value_hash (*CURRENT_SHEET*))
                   cell
                   (hash-ref (XLSX-shared_index->string_map (*XLSX*)) (string->number cell_value)))]))))

(define (to-merge-cells)
  (let ([cell_range_merge_map (SHEET-STYLE-cell_range_merge_map (*CURRENT_SHEET_STYLE*))])
    (if (> (hash-count cell_range_merge_map) 0)
        (let ([merge_count (hash-count cell_range_merge_map)])
          (append
           '("mergeCells")
           (list (cons "count" (number->string merge_count)))
           (let loop ([cell_ranges (sort (hash-keys cell_range_merge_map) string<?)]
                      [merge_list '()])
             (if (not (null? cell_ranges))
                 (loop
                  (cdr cell_ranges)
                  (cons
                   (list "mergeCell" (cons "ref" (regexp-replace* #rx"-" (car cell_ranges) ":")))
                   merge_list))
                 (reverse merge_list)))))
        '())))

(define (from-merge-cells xml_hash)
  (let ([merge_cell_range_count (hash-ref xml_hash "worksheet1.mergeCells1.mergeCell's count" 0)])
    (when (> merge_cell_range_count 0)
          (let loop-merge-cell-ranges ([loop_count 1])
            (when (<= loop_count merge_cell_range_count)
                  (hash-set! (SHEET-STYLE-cell_range_merge_map (*CURRENT_SHEET_STYLE*))
                             (hash-ref xml_hash (format "worksheet1.mergeCells1.mergeCell~a.ref" loop_count)) #t)
                  (loop-merge-cell-ranges (add1 loop_count)))))))

(define (to-col-width-style)
  (append
   (list "cols")
   (let loop ([col_range_width_list (squeeze-range-hash (SHEET-STYLE-col->width_map (*CURRENT_SHEET_STYLE*)))]
              [result_list '()])
     (if (not (null? col_range_width_list))
         (loop
          (cdr col_range_width_list)
          (cons
           (list
            "col"
            (cons "min" (number->string (first (car col_range_width_list))))
            (cons "max" (number->string (second (car col_range_width_list))))
            (cons "width" (number->string (third (car col_range_width_list)))))
           result_list))
         (reverse result_list)))))

(define (from-col-width-style xml_hash)
  (let ([col_range_width_list '()]
        [col_range_width_count (hash-ref xml_hash "worksheet1.cols1.col's count" 0)])
    (when (> col_range_width_count 0)
          (set! col_range_width_list
                (let loop-cols ([loop_count 1]
                                [result_list '()])
                  (if (<= loop_count col_range_width_count)
                      (loop-cols
                       (add1 loop_count)
                       (cons
                        (list
                         (hash-ref xml_hash (format "worksheet1.cols1.col~a.min" loop_count) "0")
                         (hash-ref xml_hash (format "worksheet1.cols1.col~a.max" loop_count) "0")
                         (hash-ref xml_hash (format "worksheet1.cols1.col~a.width" loop_count) "0"))
                        result_list))
                      (reverse result_list))))

          (let loop-col-range ([loop_list col_range_width_list])
            (when (not (null? loop_list))
                  (let ([min (string->number (first (car loop_list)))]
                        [max (string->number (second (car loop_list)))]
                        [width (string->number (third (car loop_list)))])
                    (if (or
                         (= min 0)
                         (= max 0)
                         (= width 0))
                        (loop-col-range (cdr loop_list))
                        (let loop-col ([loop_index min])
                          (when (<= loop_index max)
                                (hash-set!
                                 (SHEET-STYLE-col->width_map (*CURRENT_SHEET_STYLE*))
                                 loop_index
                                 width)
                                (loop-col (add1 loop_index))))))
                  (loop-col-range (cdr loop_list)))))))

(define (work-sheet-tail)
  '(
    ("phoneticPr" ("fontId" . "1") ("type" . "noConversion"))
    ("pageMargins" ("left" . "0.7") ("right" . "0.7") ("top" . "0.75") ("bottom" . "0.75") ("header" . "0.3") ("footer" . "0.3"))
    ("pageSetup" ("paperSize" . "9") ("orientation" . "portrait") ("horizontalDpi" . "200") ("verticalDpi" . "200") ("r:id" . "rId1"))))

(define (write-worksheets [output_dir #f])
  (let loop ([sheets (XLSX-sheet_list (*XLSX*))]
             [sheet_index 0]
             [work_sheet_index 1])

    (when (not (null? sheets))
          (if (DATA-SHEET? (car sheets))
              (let ([dir (if output_dir output_dir (build-path (XLSX-xlsx_dir (*XLSX*)) "xl" "worksheets"))])
                (make-directory* dir)
                
                (with-sheet-ref
                 sheet_index
                 (lambda ()
                   (with-output-to-file (build-path dir (format "sheet~a.xml" work_sheet_index))
                     #:exists 'replace
                     (lambda ()
                       (printf (lists->xml (to-work-sheet)))))))

                (loop (cdr sheets) (add1 sheet_index) (add1 work_sheet_index)))
              (loop (cdr sheets) (add1 sheet_index) work_sheet_index)))))

(define (read-worksheets [input_dir #f])
  (let ([dir (if input_dir input_dir (build-path (XLSX-xlsx_dir (*XLSX*)) "xl" "worksheets"))])
    (let loop ([sheets (XLSX-sheet_list (*XLSX*))]
               [sheet_index 0]
               [work_sheet_index 1])
      (when (not (null? sheets))
            (if (DATA-SHEET? (car sheets))
                (begin
                  (with-sheet-ref
                   sheet_index
                   (lambda ()
                     (from-work-sheet (xml->hash (build-path dir (format "sheet~a.xml" work_sheet_index))))))
                  (loop (cdr sheets) (add1 sheet_index) (add1 work_sheet_index)))
                (loop (cdr sheets) (add1 sheet_index) work_sheet_index))))))

