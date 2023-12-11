#lang racket

(require "../lib/dimension.rkt")
(require "../lib/sheet-lib.rkt")

(require "style-lib.rkt")
(require "style.rkt")
(require "styles.rkt")
(require "border-style.rkt")
(require "font-style.rkt")
(require "alignment-style.rkt")
(require "number-style.rkt")
(require "fill-style.rkt")

(provide (contract-out
          [set-col-range-width (-> string? natural? void?)]
          [set-row-range-height (-> string? natural? void?)]
          [set-freeze-row-col-range (-> natural? natural? void?)]
          [set-merge-cell-range (-> cell-range? void?)]
          [set-cell-range-border-style (-> string? border-direction? (or/c rgb? "") border-mode? void?)]
          [border-direction? (-> string? boolean?)]
          [set-cell-range-font-style (-> string? natural? string? rgb? void?)]
          [set-row-range-font-style (-> string? natural? string? rgb? void?)]
          [set-col-range-font-style (-> string? natural? string? rgb? void?)]
          [set-cell-range-alignment-style (-> string? horizontal_mode? vertical_mode? void?)]
          [set-row-range-alignment-style (-> string? horizontal_mode? vertical_mode? void?)]
          [set-col-range-alignment-style (-> string? horizontal_mode? vertical_mode? void?)]
          [set-cell-range-number-style (-> string? string? void?)]
          [set-row-range-number-style (-> string? string? void?)]
          [set-col-range-number-style (-> string? string? void?)]
          [set-cell-range-date-style (-> string? string? void?)]
          [set-row-range-date-style (-> string? string? void?)]
          [set-col-range-date-style (-> string? string? void?)]
          [set-cell-range-fill-style (-> string? rgb? fill-pattern? void?)]
          [set-row-range-fill-style (-> string? rgb? fill-pattern? void?)]
          [set-col-range-fill-style (-> string? rgb? fill-pattern? void?)]
          [update-style (->
                         STYLE?
                         (or/c STYLE? BORDER-STYLE? FONT-STYLE? ALIGNMENT-STYLE? NUMBER-STYLE? FILL-STYLE?)
                         STYLE?)]
          ))

(define (set-col-range-width col_range width)
  (let ([_col_range (to-col-range col_range)])
    (let loop ([col_index (car _col_range)])
      (when (<= col_index (cdr _col_range))
        (hash-set! (SHEET-STYLE-col->width_map (*CURRENT_SHEET_STYLE*)) col_index width)
        (loop (add1 col_index))))))

(define (set-row-range-height row_range height)
  (let ([_row_range (to-row-range row_range)])
    (let loop ([row_index (car _row_range)])
      (when (<= row_index (cdr _row_range))
        (hash-set! (SHEET-STYLE-row->height_map (*CURRENT_SHEET_STYLE*)) row_index height)
        (loop (add1 row_index))))))

(define (set-freeze-row-col-range rows cols)
  (set-SHEET-STYLE-freeze_range! (*CURRENT_SHEET_STYLE*) (cons rows cols)))

(define (set-merge-cell-range cell_range)
  (hash-set! (SHEET-STYLE-cell_range_merge_map (*CURRENT_SHEET_STYLE*)) cell_range #t))

(define (border-direction? direction)
  (ormap (lambda (_direction) (string=? _direction direction)) '("all" "side" "top" "bottom" "left" "right")))

(define (set-cell-range-border-style cell_range border_direction border_color border_mode)
  (cond
   [(string=? border_direction "side")
    (let-values ([(left_cells right_cells top_cells bottom_cells) (get-cell-range-four-sides-cells cell_range)])
      (add-cells-style left_cells (BORDER-STYLE border_color border_mode #f #f #f #f #f #f))
      (add-cells-style right_cells (BORDER-STYLE #f #f border_color border_mode #f #f #f #f))
      (add-cells-style top_cells (BORDER-STYLE #f #f #f #f border_color border_mode #f #f))
      (add-cells-style bottom_cells (BORDER-STYLE #f #f #f #f #f #f border_color border_mode))
      )]
   [(string=? border_direction "all")
    (add-cell-range-style cell_range (BORDER-STYLE border_color border_mode border_color border_mode border_color border_mode border_color border_mode))]
   [(string=? border_direction "left")
    (add-cell-range-style cell_range (BORDER-STYLE border_color border_mode #f #f #f #f #f #f))]
   [(string=? border_direction "right")
    (add-cell-range-style cell_range (BORDER-STYLE #f #f border_color border_mode #f #f #f #f))]
   [(string=? border_direction "top")
    (add-cell-range-style cell_range (BORDER-STYLE #f #f #f #f border_color border_mode #f #f))]
   [(string=? border_direction "bottom")
    (add-cell-range-style cell_range (BORDER-STYLE #f #f #f #f #f #f border_color border_mode))]
   ))

(define (set-cell-range-font-style cell_range font_size font_name font_color)
  (add-cell-range-style cell_range (FONT-STYLE font_size font_name font_color)))

(define (set-row-range-font-style row_range font_size font_name font_color)
  (add-row-range-style row_range (FONT-STYLE font_size font_name font_color)))

(define (set-col-range-font-style col_range font_size font_name font_color)
  (add-col-range-style col_range (FONT-STYLE font_size font_name font_color)))

(define (set-cell-range-alignment-style cell_range horizontal_placement vertical_placement)
  (add-cell-range-style cell_range (ALIGNMENT-STYLE horizontal_placement vertical_placement)))

(define (set-row-range-alignment-style row_range horizontal_placement vertical_placement)
  (add-row-range-style row_range (ALIGNMENT-STYLE horizontal_placement vertical_placement)))

(define (set-col-range-alignment-style col_range horizontal_placement vertical_placement)
  (add-col-range-style col_range (ALIGNMENT-STYLE horizontal_placement vertical_placement)))

(define (set-cell-range-number-style cell_range format)
  (add-cell-range-style cell_range (NUMBER-STYLE #f format)))

(define (set-row-range-number-style row_range format)
  (add-row-range-style row_range (NUMBER-STYLE #f format)))

(define (set-col-range-number-style col_range format)
  (add-col-range-style col_range (NUMBER-STYLE #f format)))

(define (set-cell-range-date-style cell_range format)
  (set-cell-range-number-style cell_range format))

(define (set-row-range-date-style row_range format)
  (set-row-range-number-style row_range format))

(define (set-col-range-date-style col_range format)
  (set-col-range-number-style col_range format))

(define (set-cell-range-fill-style cell_range color pattern)
  (add-cell-range-style cell_range (FILL-STYLE color pattern)))

(define (set-row-range-fill-style row_range color pattern)
  (add-row-range-style row_range (FILL-STYLE color pattern)))

(define (set-col-range-fill-style col_range color pattern)
  (add-col-range-style col_range (FILL-STYLE color pattern)))

(define (add-cell-range-style cell_range new_style)
  (add-cells-style (cell_range->cell_list cell_range) new_style))

(define (add-cells-style cells new_style)
  (let loop ([_cells cells])
    (when (not (null? _cells))
      (hash-set! (SHEET-STYLE-cell->style_map (*CURRENT_SHEET_STYLE*))
                 (car _cells)
                 (update-style
                  (hash-ref
                   (SHEET-STYLE-cell->style_map (*CURRENT_SHEET_STYLE*))
                   (car _cells) (new-style))
                  new_style))
      (loop (cdr _cells)))))

(define (add-row-range-style row_range new_style)
  (let* ([row_range (to-row-range row_range)]
         [start_row_index (car row_range)]
         [end_row_index (cdr row_range)]
         [cells (hash-keys (SHEET-STYLE-cell->style_map (*CURRENT_SHEET_STYLE*)))])

    (let loop ([loop_row_index start_row_index])
      (when (<= loop_row_index end_row_index)
        (let ([row_style (hash-ref (SHEET-STYLE-row->style_map (*CURRENT_SHEET_STYLE*)) loop_row_index (new-style))])
          (hash-set!
           (SHEET-STYLE-row->style_map (*CURRENT_SHEET_STYLE*))
           loop_row_index
           (update-style
            row_style
            new_style))
          (add-cells-style (get-row-cells loop_row_index) new_style))
        (loop (add1 loop_row_index))))))

(define (add-col-range-style col_range new_style)
  (let* ([col_range (to-col-range col_range)]
         [start_col_index (car col_range)]
         [end_col_index (cdr col_range)]
         [cells (hash-keys (SHEET-STYLE-cell->style_map (*CURRENT_SHEET_STYLE*)))])

    (let loop ([loop_col_index start_col_index])
      (when (<= loop_col_index end_col_index)
        (let ([col_style
               (hash-ref (SHEET-STYLE-col->style_map (*CURRENT_SHEET_STYLE*)) loop_col_index (new-style))])
          (hash-set!
           (SHEET-STYLE-col->style_map (*CURRENT_SHEET_STYLE*))
           loop_col_index
           (update-style col_style new_style))
          (add-cells-style (get-col-cells loop_col_index) new_style))
        (loop (add1 loop_col_index))))))

(define (update-style _style new_style)
  (let ([style (STYLE
                (if (STYLE-border_style _style)
                    (struct-copy BORDER-STYLE (STYLE-border_style _style))
                    #f)
                (if (STYLE-font_style _style)
                    (struct-copy FONT-STYLE (STYLE-font_style _style))
                    #f)
                (if (STYLE-alignment_style _style)
                    (struct-copy ALIGNMENT-STYLE (STYLE-alignment_style _style))
                    #f)
                (if (STYLE-number_style _style)
                    (struct-copy NUMBER-STYLE (STYLE-number_style _style))
                    #f)
                (if (STYLE-fill_style _style)
                    (struct-copy FILL-STYLE (STYLE-fill_style _style))
                    #f))])
    (cond
     [(BORDER-STYLE? new_style)
      (if (STYLE-border_style style)
          (update-border-style (STYLE-border_style style) new_style)
          (set-STYLE-border_style! style new_style))
      (when (not (member
                  (STYLE-border_style style)
                  (STYLES-border_list (*STYLES*))
                  equal-hash-code=?))
        (set-STYLES-border_list!
         (*STYLES*)
         `(,@(STYLES-border_list (*STYLES*)) ,(STYLE-border_style style))))
      ]
     [(FONT-STYLE? new_style)
      (if (STYLE-font_style style)
          (update-font-style (STYLE-font_style style) new_style)
          (set-STYLE-font_style! style new_style))
      (when (not (member
                  (STYLE-font_style style)
                  (STYLES-font_list (*STYLES*)) equal-hash-code=?))
        (set-STYLES-font_list!
         (*STYLES*)
         `(,@(STYLES-font_list (*STYLES*)) ,(STYLE-font_style style))))
      ]
     [(ALIGNMENT-STYLE? new_style)
      (if (STYLE-alignment_style style)
          (update-alignment-style (STYLE-alignment_style style) new_style)
          (set-STYLE-alignment_style! style new_style))]
     [(NUMBER-STYLE? new_style)
      (if (STYLE-number_style style)
          (update-number-style (STYLE-number_style style) new_style)
          (set-STYLE-number_style!
           style
           (NUMBER-STYLE
            (number->string
             (add1
              (apply
               max
               `(0
                ,@(map
                   (lambda (number_style)
                     (if (NUMBER-STYLE-formatId number_style)
                         (string->number (NUMBER-STYLE-formatId number_style))
                         0))
                   (STYLES-number_list (*STYLES*)))))))
            (NUMBER-STYLE-formatCode new_style))))

      (let ([format_exists?
             (let search-foramtId-loop ([number_styles (STYLES-number_list (*STYLES*))])
               (if (not (null? number_styles))
                   (if (string=?
                        (NUMBER-STYLE-formatCode (car number_styles))
                        (NUMBER-STYLE-formatCode (STYLE-number_style style)))
                       (begin
                         (set-NUMBER-STYLE-formatId!
                          (STYLE-number_style style)
                          (NUMBER-STYLE-formatId (car number_styles)))
                         #t)
                         (search-foramtId-loop (cdr number_styles)))
                   #f))])

        (when (not format_exists?)
          (set-NUMBER-STYLE-formatId!
           (STYLE-number_style style)
           (number->string
            (add1
             (apply
              max
              `(0
                ,@(map
                   (lambda (number_style)
                     (if (NUMBER-STYLE-formatId number_style)
                         (string->number (NUMBER-STYLE-formatId number_style))
                         0))
                   (STYLES-number_list (*STYLES*))))))))

        (set-STYLES-number_list!
         (*STYLES*)
         `(,@(STYLES-number_list (*STYLES*)) ,(STYLE-number_style style)))))
      ]
     [(FILL-STYLE? new_style)
      (if (STYLE-fill_style style)
          (update-fill-style (STYLE-fill_style style) new_style)
          (set-STYLE-fill_style! style new_style))
      (when (not (member
                  (STYLE-fill_style style)
                  (STYLES-fill_list (*STYLES*))
                  equal-hash-code=?))
        (set-STYLES-fill_list!
         (*STYLES*)
         `(,@(STYLES-fill_list (*STYLES*)) ,(STYLE-fill_style style))))
      ]
     [(STYLE? new_style)
      (set! style
            (let loop-update ([section_styles
                               (list
                                (STYLE-border_style new_style)
                                (STYLE-font_style new_style)
                                (STYLE-alignment_style new_style)
                                (STYLE-number_style new_style)
                                (STYLE-fill_style new_style))]
                              [loop_style style])
              (if (not (null? section_styles))
                  (loop-update (cdr section_styles)
                               (update-style loop_style (car section_styles)))
                  loop_style)))])

    (when (not (member style (STYLES-styles (*STYLES*)) equal-hash-code=?))
      (set-STYLES-styles! (*STYLES*) `(,@(STYLES-styles (*STYLES*)) ,style)))

    style))
