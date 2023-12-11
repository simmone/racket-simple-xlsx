#lang racket

(require "style.rkt")
(require "styles.rkt")
(require "border-style.rkt")
(require "font-style.rkt")
(require "alignment-style.rkt")
(require "number-style.rkt")
(require "fill-style.rkt")
(require "style-lib.rkt")

(provide (contract-out
          [strip-styles (-> void?)]
          [assemble-styles (-> void?)]
          ))

(define (strip-styles)
  (let ([uni_style_map (make-hash)]
        [uni_border_style_map (make-hash)]
        [uni_fill_style_map (make-hash)]
        [uni_number_style_map (make-hash)]
        [uni_font_style_map (make-hash)])

    (let loop-sheet-style ([sheet_style_list (STYLES-sheet_style_list (*STYLES*))])
      (when (not (null? sheet_style_list))
        (let ([sheet_style (car sheet_style_list)])
          (let loop ([styles
                      `(
                        ,@(hash-values (SHEET-STYLE-row->style_map sheet_style))
                        ,@(hash-values (SHEET-STYLE-col->style_map sheet_style))
                        ,@(hash-values (SHEET-STYLE-cell->style_map sheet_style)))])
            (when (not (null? styles))
              (hash-set! uni_style_map (car styles) #t)

              (when (STYLE-border_style (car styles))
                (hash-set! uni_border_style_map (STYLE-border_style (car styles)) #t))

              (when (STYLE-fill_style (car styles))
                (hash-set! uni_fill_style_map (STYLE-fill_style (car styles)) #t))

              (when (STYLE-number_style (car styles))
                (hash-set! uni_number_style_map (STYLE-number_style (car styles)) #t))

              (when (STYLE-font_style (car styles))
                (hash-set! uni_font_style_map (STYLE-font_style (car styles)) #t))

              (loop (cdr styles)))))
        (loop-sheet-style (cdr sheet_style_list))))

    (set-STYLES-styles!
     (*STYLES*)
     (filter
      (lambda (style)
        (hash-has-key? uni_style_map style))
      (STYLES-styles (*STYLES*))))

    (set-STYLES-border_list!
     (*STYLES*)
     (filter
      (lambda (border_style)
        (hash-has-key? uni_border_style_map border_style))
      (STYLES-border_list (*STYLES*))))

    (set-STYLES-font_list!
     (*STYLES*)
     (filter
      (lambda (font_style)
        (hash-has-key? uni_font_style_map font_style))
      (STYLES-font_list (*STYLES*))))

    (set-STYLES-number_list!
     (*STYLES*)
     (filter
      (lambda (number_style)
        (hash-has-key? uni_number_style_map number_style))
      (STYLES-number_list (*STYLES*))))

    (set-STYLES-fill_list!
     (*STYLES*)
     (filter
      (lambda (fill_style)
        (hash-has-key? uni_fill_style_map fill_style))
      (STYLES-fill_list (*STYLES*))))
    ))

(define (assemble-styles)
  (set-STYLES-styles!
   (*STYLES*)
   (append
    (list (STYLE #f #f #f #f (FILL-STYLE "FFFFFF" "none"))
          (STYLE #f #f #f #f (FILL-STYLE "FFFFFF" "gray125")))
    (remove*
     (list (STYLE #f #f #f #f (FILL-STYLE "FFFFFF" "none"))
           (STYLE #f #f #f #f (FILL-STYLE "FFFFFF" "gray125")))
     (STYLES-styles (*STYLES*))
     equal-hash-code=?)))

  (set-STYLES-border_list!
   (*STYLES*)
   (append
    (list
     (BORDER-STYLE #f #f #f #f #f #f #f #f)
     )
    (remove*
     (list
      (BORDER-STYLE #f #f #f #f #f #f #f #f)
      )
     (STYLES-border_list (*STYLES*))
     equal-hash-code=?)))

  (set-STYLES-font_list!
   (*STYLES*)
   (append
    (list
     (FONT-STYLE 10 "Arial" "000000")
     )
    (remove*
     (list
      (FONT-STYLE 10 "Arial" "000000")
      )
     (STYLES-font_list (*STYLES*))
     equal-hash-code=?)))

  (set-STYLES-fill_list!
   (*STYLES*)
   (append
    (list (FILL-STYLE "FFFFFF" "none")
          (FILL-STYLE "FFFFFF" "gray125"))
    (remove*
     (list (FILL-STYLE "FFFFFF" "none")
           (FILL-STYLE "FFFFFF" "gray125"))
     (STYLES-fill_list (*STYLES*))
     equal-hash-code=?)))
)
