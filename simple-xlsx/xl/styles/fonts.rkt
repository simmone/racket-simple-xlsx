#lang racket

(require "../../style/font-style.rkt")

(provide (contract-out
          [to-fonts (-> (listof FONT-STYLE?) list?)]
          [from-fonts (-> hash? void?)]
          ))

(define (to-fonts font_list)
  (append
   (list "fonts" (cons "count" (number->string (add1 (length font_list)))))
   '(("font"
      ("sz" ("val" . "11"))
      ("color" ("theme" . "1"))
      ("name" ("val" . "Arial"))
      ("family" ("val" . "2"))
      ("charset" ("val" . "134"))
      ("scheme" ("val" . "minor"))))
   (let loop ([fonts font_list]
              [result_list '()])
     (if (not (null? fonts))
         (loop
          (cdr fonts)
          (cons
           (list
            "font"
            (list "sz" (cons "val" (number->string (FONT-STYLE-size (car fonts)))))
            (if (FONT-STYLE-color (car fonts))
                (list "color" (cons "rgb" (FONT-STYLE-color (car fonts))))
                (list "color" (cons "theme" "1")))
            (list "name" (cons "val" (FONT-STYLE-name (car fonts))))
            '("family" ("val" . "2"))
            (if (not (regexp-match #rx"^([a-zA-Z]| |-|_|[0-9])+$" (FONT-STYLE-name (car fonts))))
                (list "charset" (cons "val" "134"))
                '())
            (list "scheme" (cons "val" "minor")))
           result_list))
         (reverse result_list)))))

(define (from-fonts xml_hash)
  (let ([skip_count 1])
    (let loop ([loop_count 0])
      (when (< loop_count (hash-ref xml_hash "styleSheet1.fonts1.font's count" 0))
            (let ([prefix (format "styleSheet1.fonts1.font~a" (add1 loop_count))])
              (when (>= loop_count skip_count)
                    (let ([font_style_hash_code
                           (FONT-STYLE-hash_code
                            (FONT-STYLE
                             ""
                             (string->number (hash-ref xml_hash (format "~a.sz1.val" prefix)))
                             (hash-ref xml_hash (format "~a.name1.val" prefix))
                             (hash-ref xml_hash (format "~a.color1.rgb" prefix) "000000")))])

                      (hash-set! (*FONT_STYLE->INDEX_MAP*) font_style_hash_code (- loop_count skip_count))
                      (hash-set! (*FONT_INDEX->STYLE_MAP*) (- loop_count skip_count) font_style_hash_code))))
            (loop (add1 loop_count))))))
