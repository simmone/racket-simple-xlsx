#lang racket

(require "../lib/dimension.rkt"
         "style-lib.rkt"
         "border-style.rkt"
         "font-style.rkt"
         "number-style.rkt"
         "fill-style.rkt"
         "style.rkt")

(provide (contract-out
          [struct STYLES
                  (
                   (styles (listof STYLE?))
                   (border_list (listof BORDER-STYLE?))
                   (font_list (listof FONT-STYLE?))
                   (number_list (listof NUMBER-STYLE?))
                   (fill_list (listof FILL-STYLE?))
                   (sheet_style_list (listof SHEET-STYLE?))
                   )]
          [new-styles (-> STYLES?)]
          [*STYLES* (parameter/c STYLES?)]
          [struct SHEET-STYLE
                  (
                   (cell->style_map (hash/c string? STYLE?))
                   (row->style_map (hash/c natural? STYLE?))
                   (col->style_map (hash/c natural? STYLE?))
                   (col->width_map (hash/c natural? number?))
                   (row->height_map (hash/c natural? number?))
                   (freeze_range (cons/c natural? natural?))
                   (cell_range_merge_map (hash/c cell-range? #t))
                   )
                  ]
          [new-sheet-style (-> SHEET-STYLE?)]
          [*CURRENT_SHEET_STYLE* (parameter/c SHEET-STYLE?)]
          ))

(define *STYLES* (make-parameter #f))
(define *CURRENT_SHEET_STYLE* (make-parameter #f))

(struct STYLES
        (
         (styles #:mutable)
         (border_list #:mutable)
         (font_list #:mutable)
         (number_list #:mutable)
         (fill_list #:mutable)
         (sheet_style_list #:mutable)
         )
        #:transparent
        )

(define (new-styles)
  (
   STYLES
   '()
   '()
   '()
   '()
   '()
   '()))

(struct SHEET-STYLE
        (
         cell->style_map
         row->style_map
         col->style_map
         col->width_map
         row->height_map
         (freeze_range #:mutable)
         cell_range_merge_map
         )
        )

(define (new-sheet-style)
  (SHEET-STYLE
   (make-hash)
   (make-hash)
   (make-hash)
   (make-hash)
   (make-hash)
   '(0 . 0)
   (make-hash)
   )
  )
