#lang racket

(require "../lib/dimension.rkt")

(require "lib.rkt")
(require "border-style.rkt")
(require "font-style.rkt")
(require "alignment-style.rkt")
(require "number-style.rkt")
(require "fill-style.rkt")

(provide (contract-out
          [struct STYLES
                  (
                   (style->index_map (hash/c string? natural?))
                   (index->style_map (hash/c natural? string?))

                   (border_style->index_map (hash/c string? natural?))
                   (border_index->style_map (hash/c natural? string?))

                   (font_style->index_map (hash/c string? natural?))
                   (font_index->style_map (hash/c natural? string?))

                   (number_style->index_map (hash/c string? natural?))
                   (number_index->style_map (hash/c natural? string?))

                   (fill_style->index_map (hash/c string? natural?))
                   (fill_index->style_map (hash/c natural? string?))

                   (sheet_index->style_map (hash/c natural? SHEET-STYLE?))
                   )]
          [new-styles (-> STYLES?)]
          [*STYLES* (parameter/c (or/c STYLES? #f))]
          [*STYLE->INDEX_MAP* (parameter/c (or/c (hash/c string? natural?) #f))]
          [*INDEX->STYLE_MAP* (parameter/c (or/c (hash/c natural? string?) #f))]
          [*CURRENT_SHEET_STYLE* (parameter/c (or/c SHEET-STYLE? #f))]
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
          [struct STYLE
                  (
                   (hash_code string?)
                   (border_style (or/c #f BORDER-STYLE?))
                   (font_style (or/c #f FONT-STYLE?))
                   (alignment_style (or/c #f ALIGNMENT-STYLE?))
                   (number_style (or/c #f NUMBER-STYLE?))
                   (fill_style (or/c #f FILL-STYLE?))
                   )]
          [new-style (-> STYLE?)]
          [style-null? (-> STYLE? boolean?)]
          [style-from-hash-code (-> string? STYLE?)]
          ))

(define *STYLES* (make-parameter #f))
(define *STYLE->INDEX_MAP* (make-parameter #f))
(define *INDEX->STYLE_MAP* (make-parameter #f))
(define *CURRENT_SHEET_STYLE* (make-parameter #f))

(struct STYLES
        (
         (style->index_map #:mutable)
         (index->style_map #:mutable)
         (border_style->index_map #:mutable)
         (border_index->style_map #:mutable)
         (font_style->index_map #:mutable)
         (font_index->style_map #:mutable)
         (number_style->index_map #:mutable)
         (number_index->style_map #:mutable)
         (fill_style->index_map #:mutable)
         (fill_index->style_map #:mutable)
         (sheet_index->style_map #:mutable)
         ))
(define (new-styles) (STYLES
                      (make-hash)
                      (make-hash)

                      (make-hash)
                      (make-hash)

                      (make-hash)
                      (make-hash)

                      (make-hash)
                      (make-hash)

                      (make-hash)
                      (make-hash)

                      (make-hash)))

(struct SHEET-STYLE
        (cell->style_map row->style_map col->style_map col->width_map row->height_map (freeze_range #:mutable) cell_range_merge_map
        ))
(define (new-sheet-style) (SHEET-STYLE (make-hash) (make-hash) (make-hash) (make-hash) (make-hash) '(0 . 0) (make-hash)))

(struct STYLE
        (hash_code
         (border_style #:mutable)
         (font_style #:mutable)
         (alignment_style #:mutable)
         (number_style #:mutable)
         (fill_style #:mutable)
         )
        #:transparent
        #:guard
        (lambda (_hash_code _border_style _font_style _alignment_style _number_style _fill_style name)
          (values
           (format "~a<s>~a<s>~a<s>~a<s>~a"
                   (if _border_style (BORDER-STYLE-hash_code _border_style) "")
                   (if _font_style (FONT-STYLE-hash_code _font_style) "")
                   (if _alignment_style (ALIGNMENT-STYLE-hash_code _alignment_style) "")
                   (if _number_style (NUMBER-STYLE-hash_code _number_style) "")
                   (if _fill_style (FILL-STYLE-hash_code _fill_style) "")
                   )
           _border_style _font_style _alignment_style _number_style _fill_style)))
(define (new-style) (STYLE "" #f #f #f #f #f))

(define (style-null? style)
  (if (string=? (STYLE-hash_code style) (STYLE-hash_code (new-style)))
      #t
      #f))

(define (style-from-hash-code hash_code)
  (let ([items (regexp-split #rx"<s>" hash_code)])
    (if (= (length items) 5)
        (STYLE ""
         (border-style-from-hash-code (list-ref items 0))
         (font-style-from-hash-code (list-ref items 1))
         (alignment-style-from-hash-code (list-ref items 2))
         (number-style-from-hash-code (list-ref items 3))
         (fill-style-from-hash-code (list-ref items 4)))
        (new-style))))
