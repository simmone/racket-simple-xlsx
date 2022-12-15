#lang racket

(require "style.rkt")
(require "border-style.rkt")
(require "font-style.rkt")
(require "alignment-style.rkt")
(require "number-style.rkt")
(require "fill-style.rkt")

(provide (contract-out
          [style<? (-> STYLE? STYLE? boolean?)]
          [sort-styles (-> void?)]
          ))

(define (style<? style1 style2)
  (cond
   [(string=? (STYLE-hash_code style1) (STYLE-hash_code style2))
      #f]
   [(not (border-style=? (STYLE-border_style style1) (STYLE-border_style style2)))
    (border-style<? (STYLE-border_style style1) (STYLE-border_style style2))]
   [(not (font-style=? (STYLE-font_style style1) (STYLE-font_style style2)))
    (font-style<? (STYLE-font_style style1) (STYLE-font_style style2))]
   [(not (alignment-style=? (STYLE-alignment_style style1) (STYLE-alignment_style style2)))
    (alignment-style<? (STYLE-alignment_style style1) (STYLE-alignment_style style2))]
   [(not (number-style=? (STYLE-number_style style1) (STYLE-number_style style2)))
    (number-style<? (STYLE-number_style style1) (STYLE-number_style style2))]
   [(not (fill-style=? (STYLE-fill_style style1) (STYLE-fill_style style2)))
    (fill-style<? (STYLE-fill_style style1) (STYLE-fill_style style2))]
   [else
    #f]))

(define (sort-styles)
  (let loop-sheet-style ([sheet_styles (hash-values (STYLES-sheet_index->style_map (*STYLES*)))])
    (when (not (null? sheet_styles))
          (let loop ([styles
                      `(
                        ,@(hash-values (SHEET-STYLE-row->style_map (car sheet_styles)))
                        ,@(hash-values (SHEET-STYLE-col->style_map (car sheet_styles)))
                        ,@(hash-values (SHEET-STYLE-cell->style_map (car sheet_styles))))])
            (when (not (null? styles))
              (when (not (style-null? (car styles)))
                (hash-set! (STYLES-style->index_map (*STYLES*)) (STYLE-hash_code (car styles)) 0)
                (when (STYLE-border_style (car styles))
                  (hash-set!
                   (STYLES-border_style->index_map (*STYLES*))
                   (BORDER-STYLE-hash_code (STYLE-border_style (car styles))) 0))
                (when (STYLE-font_style (car styles))
                  (hash-set!
                   (STYLES-font_style->index_map (*STYLES*))
                   (FONT-STYLE-hash_code (STYLE-font_style (car styles))) 0))
                (when (STYLE-number_style (car styles))
                  (hash-set!
                   (STYLES-number_style->index_map (*STYLES*))
                   (NUMBER-STYLE-hash_code (STYLE-number_style (car styles))) 0))
                (when (STYLE-fill_style (car styles))
                  (hash-set!
                   (STYLES-fill_style->index_map (*STYLES*))
                   (FILL-STYLE-hash_code (STYLE-fill_style (car styles))) 0)))
              (loop (cdr styles))))
          (loop-sheet-style (cdr sheet_styles))))

  (hash-clear! (STYLES-index->style_map (*STYLES*)))
  (let loop ([sorted_style_list (sort
                                 (hash->list (STYLES-style->index_map (*STYLES*)))
                                 #:key
                                 (lambda (styles)
                                   (style-from-hash-code (car styles)))
                                 style<?)]
             [loop_index 0])
    (when (not (null? sorted_style_list))
      (hash-set! (STYLES-style->index_map (*STYLES*)) (caar sorted_style_list) loop_index)
      (hash-set! (STYLES-index->style_map (*STYLES*)) loop_index (caar sorted_style_list))
      (loop (cdr sorted_style_list) (add1 loop_index))))

  (let loop ([sorted_border_style_list (sort (hash->list (STYLES-border_style->index_map (*STYLES*))) string<? #:key car)]
             [loop_index 0])
    (when (not (null? sorted_border_style_list))
      (hash-set! (STYLES-border_style->index_map (*STYLES*)) (caar sorted_border_style_list) loop_index)
      (hash-set! (STYLES-border_index->style_map (*STYLES*)) loop_index (caar sorted_border_style_list))
      (loop (cdr sorted_border_style_list) (add1 loop_index))))

  (let loop ([sorted_font_style_list (sort (hash->list (STYLES-font_style->index_map (*STYLES*))) string<? #:key car)]
             [loop_index 0])
    (when (not (null? sorted_font_style_list))
      (hash-set! (STYLES-font_style->index_map (*STYLES*)) (caar sorted_font_style_list) loop_index)
      (hash-set! (STYLES-font_index->style_map (*STYLES*)) loop_index (caar sorted_font_style_list))
      (loop (cdr sorted_font_style_list) (add1 loop_index))))

  (let loop ([sorted_number_style_list (sort (hash->list (STYLES-number_style->index_map (*STYLES*))) string<? #:key car)]
             [loop_index 0])
    (when (not (null? sorted_number_style_list))
      (hash-set! (STYLES-number_style->index_map (*STYLES*)) (caar sorted_number_style_list) loop_index)
      (hash-set! (STYLES-number_index->style_map (*STYLES*)) loop_index (caar sorted_number_style_list))
      (loop (cdr sorted_number_style_list) (add1 loop_index))))

  (let loop ([sorted_fill_style_list (sort (hash->list (STYLES-fill_style->index_map (*STYLES*))) string<? #:key car)]
             [loop_index 0])
    (when (not (null? sorted_fill_style_list))
      (hash-set! (STYLES-fill_style->index_map (*STYLES*)) (caar sorted_fill_style_list) loop_index)
      (hash-set! (STYLES-fill_index->style_map (*STYLES*)) loop_index (caar sorted_fill_style_list))
      (loop (cdr sorted_fill_style_list) (add1 loop_index))))
  )
