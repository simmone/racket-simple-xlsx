#lang racket

(require "../lib/dimension.rkt")

(provide (contract-out
          [cell-value? (-> (or/c string? number? date?) boolean?)]
          [struct DATA-SHEET
                  (
                   (sheet_name string?)
                   (dimension cell-range?)
                   (cell->value_hash (hash/c cell? cell-value?))
                   )]
          [struct CHART-SHEET
                  (
                   (sheet_name string?)
                   (chart_type (or/c 'LINE 'LINE3D 'BAR 'BAR3D 'PIE 'PIE3D 'UNKNOWN))
                   (topic string?)
                   (serial (listof (list/c string? string? cell-range? string? cell-range?)))
                   )]
          [get-sheet-name (-> (or/c DATA-SHEET? CHART-SHEET?) string?)]
          ))

(struct DATA-SHEET (
                    [sheet_name #:mutable]
                    [dimension #:mutable]
                    [cell->value_hash #:mutable]
                    ))

(define (cell-value? val)
  (if (or
       (string? val)
       (number? val)
       (date? val))
      #t
      #f))

(struct CHART-SHEET (
                     [sheet_name #:mutable]
                     [chart_type #:mutable]
                     [topic #:mutable]
                     [serial #:mutable]))

(define (get-sheet-name sheet)
  (cond
   [(DATA-SHEET? sheet)
    (DATA-SHEET-sheet_name sheet)]
   [(CHART-SHEET? sheet)
    (CHART-SHEET-sheet_name sheet)]
   [else
    ""]))
