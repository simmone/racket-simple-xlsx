#lang racket

(require "../../style/font-style.rkt")
(require "../../style/styles.rkt")

(provide (contract-out
          [to-fonts (-> (listof FONT-STYLE?) list?)]
          [from-fonts (-> hash? void?)]
          ))

(define (to-fonts font_list)
  (append
   (list "fonts" (cons "count" (number->string (length font_list))))
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
            (list "name" (cons "val" (FONT-STYLE-name (car fonts)))))
           result_list))
         (reverse result_list)))))

(define (from-fonts xml_hash)
  (let loop ([loop_count 0]
             [font_list '()])
    (if (< loop_count (hash-ref xml_hash "styleSheet1.fonts1.font's count" 0))
        (let* ([prefix (format "styleSheet1.fonts1.font~a" (add1 loop_count))]
               [font_style
                 (FONT-STYLE
                  (string->number (hash-ref xml_hash (format "~a.sz1.val" prefix) "10"))
                  (hash-ref xml_hash (format "~a.name1.val" prefix))
                  (hash-ref xml_hash (format "~a.color1.rgb" prefix) "000000"))])
            (loop (add1 loop_count) (cons font_style font_list)))
        (set-STYLES-font_list! (*STYLES*) (reverse font_list)))))
