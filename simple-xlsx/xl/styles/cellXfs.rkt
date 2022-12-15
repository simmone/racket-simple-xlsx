#lang racket

(require "../../style/style.rkt")
(require "../../style/border-style.rkt")
(require "../../style/fill-style.rkt")
(require "../../style/alignment-style.rkt")
(require "../../style/number-style.rkt")
(require "../../style/font-style.rkt")

(provide (contract-out
          [to-cellXfs (-> (listof STYLE?) list?)]
          [from-cellXfs (-> hash? void?)]
          ))

(define SKIP_BORDER_COUNT 1)
(define SKIP_FONT_COUNT 1)
(define SKIP_NUMBER_COUNT 1)
(define SKIP_FILL_COUNT 2)

(define (to-cellXfs style_list)
  (append
   (list "cellXfs" (cons "count" (number->string (add1 (length style_list)))))
   '(("xf"
      ("borderId" . "0") ("fontId" . "0") ("numFmtId" . "0") ("fillId" . "0") ("xfId" . "0")
      ("alignment" ("vertical" . "center"))))
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
                  (cons "borderId"
                        (if (STYLE-border_style (car styles))
                            (number->string
                             (+
                              SKIP_BORDER_COUNT
                              (hash-ref
                               (STYLES-border_style->index_map (*STYLES*))
                               (BORDER-STYLE-hash_code (STYLE-border_style (car styles))))))
                            "0"))
                  (cons "fontId"
                        (if (STYLE-font_style (car styles))
                            (number->string
                             (+
                              SKIP_FONT_COUNT
                              (hash-ref
                               (STYLES-font_style->index_map (*STYLES*))
                               (FONT-STYLE-hash_code (STYLE-font_style (car styles))))))
                            "0"))
                  (cons "numFmtId"
                        (if (STYLE-number_style (car styles))
                            (number->string
                             (+
                              SKIP_NUMBER_COUNT
                              164
                              (hash-ref
                               (STYLES-number_style->index_map (*STYLES*))
                               (NUMBER-STYLE-hash_code (STYLE-number_style (car styles))))))
                            "0"))
                  (cons "fillId"
                        (if (STYLE-fill_style (car styles))
                            (number->string
                             (+
                              SKIP_FILL_COUNT
                              (hash-ref
                               (STYLES-fill_style->index_map (*STYLES*))
                               (FILL-STYLE-hash_code (STYLE-fill_style (car styles))))))
                            "0"))
                  (cons "xfId" "0"))
            (if (STYLE-border_style (car styles)) '(("applyBorder" . "1")) '())
            (if (STYLE-font_style (car styles)) '(("applyFont" . "1")) '())
            (if (STYLE-fill_style (car styles)) '(("applyFill" . "1")) '())
            (if (STYLE-alignment_style (car styles))
                (list
                 (list "alignment"
                       (cons "horizontal" (ALIGNMENT-STYLE-horizontal_placement (STYLE-alignment_style (car styles))))
                       (cons "vertical" (ALIGNMENT-STYLE-vertical_placement (STYLE-alignment_style (car styles))))))
                '(("alignment" ("horizontal" . "left") ("vertical" . "center")))))
           result_list))
         (reverse result_list)))))

(define (from-cellXfs xml_hash)
  (let loop ([loop_count 0])
    (when (< loop_count (hash-ref xml_hash "styleSheet1.cellXfs1.xf's count" 0))
          (when (>= loop_count 1)
                (let ([prefix (format "styleSheet1.cellXfs1.xf~a" (add1 loop_count))]
                      [border_style #f]
                      [font_style #f]
                      [alignment_style #f]
                      [number_style #f]
                      [fill_style #f])

                  (when (hash-has-key? xml_hash (format "~a.applyFill" prefix))
                        (set! fill_style
                              (fill-style-from-hash-code
                               (hash-ref (*FILL_INDEX->STYLE_MAP*)
                                         (-
                                          (string->number (hash-ref xml_hash (format "~a.fillId" prefix)))
                                          SKIP_FILL_COUNT)))))

                  (when (hash-has-key? xml_hash (format "~a.applyFont" prefix))
                        (set! font_style
                              (font-style-from-hash-code
                               (hash-ref (*FONT_INDEX->STYLE_MAP*)
                                         (-
                                          (string->number (hash-ref xml_hash (format "~a.fontId" prefix)))
                                          SKIP_FONT_COUNT)))))

                  (when (hash-has-key? xml_hash (format "~a.applyBorder" prefix))
                        (set! border_style
                              (border-style-from-hash-code
                               (hash-ref (*BORDER_INDEX->STYLE_MAP*)
                                         (-
                                          (string->number (hash-ref xml_hash (format "~a.borderId" prefix)))
                                          SKIP_BORDER_COUNT)))))

                  (when
                   (and
                    (hash-has-key? xml_hash (format "~a.numFmtId" prefix))
                    (> (string->number (hash-ref xml_hash (format "~a.numFmtId" prefix))) 0))
                   (set! number_style
                         (number-style-from-hash-code
                          (hash-ref (*NUMBER_INDEX->STYLE_MAP*)
                                    (-
                                     (string->number (hash-ref xml_hash (format "~a.numFmtId" prefix)))
                                     SKIP_NUMBER_COUNT
                                     164)))))

                  (set! alignment_style (ALIGNMENT-STYLE
                                         ""
                                         (hash-ref xml_hash (format "~a.alignment1.horizontal" prefix) "center")
                                         (hash-ref xml_hash (format "~a.alignment1.vertical" prefix) "center")))

                  (let ([style_hash_code
                         (STYLE-hash_code
                          (STYLE
                           ""
                           border_style
                           font_style
                           alignment_style
                           number_style
                           fill_style))])

                    (hash-set! (*STYLE->INDEX_MAP*) style_hash_code loop_count)
                    (hash-set! (*INDEX->STYLE_MAP*) loop_count style_hash_code))))
          (loop (add1 loop_count)))))
