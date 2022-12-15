#lang racket

(require "../../style/border-style.rkt")

(provide (contract-out
          [to-borders (-> (listof BORDER-STYLE?) list?)]
          [from-borders (-> hash? void?)]
          ))

(define (to-borders border_list)
  (append
   '("borders")
   (list (cons "count" (number->string (add1 (length border_list)))))
   '(("border" ("left") ("right") ("top") ("bottom") ("diagonal")))
   (let loop ([borders border_list]
              [result_list '()])
     (if (not (null? borders))
         (loop
          (cdr borders)
          (cons
           (append
            '("border")
            (if (not (string=? (BORDER-STYLE-left_mode (car borders)) ""))
                (list
                 (list
                  "left"
                  (cons "style" (BORDER-STYLE-left_mode (car borders)))
                  (list "color" (cons "rgb" (BORDER-STYLE-left_color (car borders))))))
                '())
            (if (not (string=? (BORDER-STYLE-right_mode (car borders)) ""))
                (list
                 (list
                  "right"
                  (cons "style" (BORDER-STYLE-right_mode (car borders)))
                  (list "color" (cons "rgb" (BORDER-STYLE-right_color (car borders))))))
                '())
            (if (not (string=? (BORDER-STYLE-top_mode (car borders)) ""))
                (list
                 (list
                  "top"
                  (cons "style" (BORDER-STYLE-top_mode (car borders)))
                  (list "color" (cons "rgb" (BORDER-STYLE-top_color (car borders))))))
                '())
            (if (not (string=? (BORDER-STYLE-bottom_mode (car borders)) ""))
                (list
                 (list
                  "bottom"
                  (cons "style" (BORDER-STYLE-bottom_mode (car borders)))
                  (list "color" (cons "rgb" (BORDER-STYLE-bottom_color (car borders))))))
                '())
            '(("diagonal")))
           result_list))
         (reverse result_list)))))

(define (from-borders xml_hash)
  (let ([skip_count 1])
    (let loop ([loop_count 0])
      (when (< loop_count (hash-ref xml_hash "styleSheet1.borders1.border's count" 0))
            (when (>= loop_count skip_count)
                  (let* ([prefix (format "styleSheet1.borders1.border~a" (add1 loop_count))]
                         [hash_code
                          (BORDER-STYLE-hash_code
                           (BORDER-STYLE
                            ""
                            (hash-ref xml_hash (format "~a.top1.color1.rgb" prefix) "")
                            (hash-ref xml_hash (format "~a.top1.style" prefix) "")
                            (hash-ref xml_hash (format "~a.bottom1.color1.rgb" prefix) "")
                            (hash-ref xml_hash (format "~a.bottom1.style" prefix) "")
                            (hash-ref xml_hash (format "~a.left1.color1.rgb" prefix) "")
                            (hash-ref xml_hash (format "~a.left1.style" prefix) "")
                            (hash-ref xml_hash (format "~a.right1.color1.rgb" prefix) "")
                            (hash-ref xml_hash (format "~a.right1.style" prefix) "")))])

                    (hash-set! (*BORDER_STYLE->INDEX_MAP*) hash_code (- loop_count skip_count))
                    (hash-set! (*BORDER_INDEX->STYLE_MAP*) (- loop_count skip_count) hash_code)))
        (loop (add1 loop_count))))))

