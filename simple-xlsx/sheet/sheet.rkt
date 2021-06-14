#lang racket

(provide (contract-out
          [struct DATA-SHEET
                  (
                   (dimension (cons/c natural? natural?))
                   (rvtsf_map (hash/c string? (list/c (or/c string? #f) (or/c string? #f) (or/c string? #f) (or/c string? #f))))
                   (width_hash hash?)
                   (height_hash hash?)
                   (freeze_range (cons/c natural? natural?))
                   (cell->style_index_map hash?)
                   (row->style_index_map hash?)
                   (col->style_index_map hash?)
                   )]
          [new-data-sheet (->
                           (cons/c natural? natural?)
                           (hash/c string? (list/c (or/c string? #f) (or/c string? #f) (or/c string? #f) (or/c string? #f)))
                           DATA-SHEET?)]
          [struct CHART-SHEET
                  (
                   (chart_type (or/c 'LINE 'LINE3D 'BAR 'BAR3D 'PIE 'PIE3D))
                   (topic string?)
                   (x_topic string?)
                   (ref_sheet_name string?)
                   (ref_range string?)
                   (y_data_range_list list?)
                   )]
          [new-chart-sheet (-> (or/c 'LINE 'LINE3D 'BAR 'BAR3D 'PIE 'PIE3D) string? CHART-SHEET?)]
          [struct DATA-RANGE
                  (
                   (sheet_name string?)
                   (range_str string?)
                   )]
          [*CURRENT_SHEET* (parameter/c (or/c DATA-SHEET? CHART-SHEET? #f))]
          ))

(define *CURRENT_SHEET* (make-parameter #f))

(struct DATA-SHEET (
                    [dimension #:mutable]
                    [rvtsf_map #:mutable]
                    [width_hash #:mutable]
                    [height_hash #:mutable]
                    [freeze_range #:mutable]
                    [cell->style_index_map #:mutable]
                    [row->style_index_map #:mutable]
                    [col->style_index_map #:mutable]
                    ))

(define (new-data-sheet dimension rvtsf_map)
  (DATA-SHEET
   dimension
   rvtsf_map
   (make-hash) (make-hash) '(0 . 0)
   (make-hash) (make-hash) (make-hash)))

(struct COL-ATTR ([width #:mutable] [back_color #:mutable]))

(struct CHART-SHEET (
                     [chart_type #:mutable] 
                     [topic #:mutable] 
                     [x_topic #:mutable] 
                     [ref_sheet_name #:mutable]
                     [ref_range #:mutable]
                     [y_data_range_list #:mutable]))

(struct DATA-RANGE ([sheet_name #:mutable] [range_str #:mutable]))

(define (new-chart-sheet chart_type topic)
  (CHART-SHEET
   chart_type topic "" "" "" '()))
