#lang racket

(require "../../style/style.rkt"
         "../../style/style-lib.rkt"
         "../../style/styles.rkt"
         "../../style/border-style.rkt"
         "../../style/fill-style.rkt"
         "../../style/alignment-style.rkt"
         "../../style/number-style.rkt"
         "../../style/font-style.rkt")

(provide (contract-out
          [to-cellXfs (-> (listof STYLE?) list?)]
          [from-cellXfs (-> hash? void?)]
          ))

(define (to-cellXfs style_list)
  (append
   (list "cellXfs" (cons "count" (number->string (length style_list))))
   (let loop ([loop_count 0]
              [styles style_list]
              [result_list '()])
     (if (not (null? styles))
         (loop
          (add1 loop_count)
          (cdr styles)
          (cons
           (append
            (list "xf"
                  (if (STYLE-border_style (car styles))
                      (cons "borderId"
                            (number->string
                             (index-of
                              (STYLES-border_list (*STYLES*))
                              (STYLE-border_style (car styles))
                              equal-hash-code=?)))
                      '())
                  (if (STYLE-font_style (car styles))
                      (cons "fontId"
                            (number->string
                             (index-of
                              (STYLES-font_list (*STYLES*))
                              (STYLE-font_style (car styles))
                              equal-hash-code=?)))
                      '())
                  (if (STYLE-number_style (car styles))
                      (cons "numFmtId" (NUMBER-STYLE-formatId (STYLE-number_style (car styles))))
                      '())
                  (if (STYLE-fill_style (car styles))
                      (cons "fillId"
                            (number->string
                             (index-of
                              (STYLES-fill_list (*STYLES*))
                              (STYLE-fill_style (car styles))
                              equal-hash-code=?)))
                      '())
                  (cons "xfId" "0"))
            (if (STYLE-border_style (car styles)) '(("applyBorder" . "1")) '())
            (if (STYLE-font_style (car styles)) '(("applyFont" . "1")) '())
            (if (STYLE-fill_style (car styles)) '(("applyFill" . "1")) '())
            (if (STYLE-alignment_style (car styles))
                (list
                 (list "alignment"
                       (cons "horizontal"
                             (let ([horizontal_placement (ALIGNMENT-STYLE-horizontal_placement (STYLE-alignment_style (car styles)))])
                               (if (string=? horizontal_placement "") "center" horizontal_placement)))
                       (cons "vertical"
                             (let ([vertical_placement (ALIGNMENT-STYLE-vertical_placement (STYLE-alignment_style (car styles)))])
                               (if (string=? vertical_placement "") "bottom" vertical_placement)))))
                '(("alignment" ("horizontal" . "center") ("vertical" . "bottom")))))
           result_list))
         (reverse result_list)))))

(define (from-cellXfs xml_hash)
  (let loop ([loop_count 0]
             [style_list '()])
    (if (< loop_count (hash-ref xml_hash "styleSheet1.cellXfs1.xf's count" 0))
        (let ([prefix (format "styleSheet1.cellXfs1.xf~a" (add1 loop_count))]
              [border_style #f]
              [font_style #f]
              [alignment_style #f]
              [number_style #f]
              [fill_style #f])

          ;; appFill is not mandatory
          (when (hash-has-key? xml_hash (format "~a.fillId1" prefix))
            (set! fill_style
                  (list-ref (STYLES-fill_list (*STYLES*))
                            (string->number (hash-ref xml_hash (format "~a.fillId1" prefix))))))

          (when (hash-has-key? xml_hash (format "~a.applyFont1" prefix))
            (set! font_style
                  (list-ref (STYLES-font_list (*STYLES*))
                            (string->number (hash-ref xml_hash (format "~a.fontId1" prefix))))))

          (when (hash-has-key? xml_hash (format "~a.applyBorder1" prefix))
            (set! border_style
                  (list-ref (STYLES-border_list (*STYLES*))
                            (string->number (hash-ref xml_hash (format "~a.borderId1" prefix))))))

          ;; Some App like Google Sheets has its own standard number formats,
          ;; So you can't skip numFmtId even it is not in the numFmts.
          (when
              (and
               (hash-has-key? xml_hash (format "~a.numFmtId1" prefix))
               (> (string->number (hash-ref xml_hash (format "~a.numFmtId1" prefix))) 0))
            (set! number_style
                  (let ([number_fmt_id (hash-ref xml_hash (format "~a.numFmtId1" prefix))])
                    (let search-loop ([number_styles (STYLES-number_list (*STYLES*))])
                      (if (not (null? number_styles))
                          (if (string=? number_fmt_id (NUMBER-STYLE-formatId (car number_styles)))
                              (car number_styles)
                              (search-loop (cdr number_styles)))
                          (NUMBER-STYLE number_fmt_id 'APP))))))

          (set! alignment_style (ALIGNMENT-STYLE
                                 (hash-ref xml_hash (format "~a.alignment1.horizontal1" prefix) "center")
                                 (hash-ref xml_hash (format "~a.alignment1.vertical1" prefix) "bottom")))

          (let ([style
                    (STYLE
                     border_style
                     font_style
                     alignment_style
                     number_style
                     fill_style)])
            (loop (add1 loop_count) (cons style style_list))))
        (set-STYLES-styles! (*STYLES*) (reverse style_list)))))
