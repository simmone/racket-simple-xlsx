#lang racket

(require "../../style/styles.rkt"
         "../../style/border-style.rkt")

(provide (contract-out
          [to-borders (-> (listof BORDER-STYLE?) list?)]
          [from-borders (-> hash? void?)]
          ))

(define (to-borders border_list)
  (append
   '("borders")
   (list (cons "count" (number->string (length border_list))))
   (let loop ([borders border_list]
              [result_list '()])
     (if (not (null? borders))
         (loop
          (cdr borders)
          (cons
           (append
            '("border")
            (if (BORDER-STYLE-left_mode (car borders))
                (list
                 (list
                  "left"
                  (cons "style" (BORDER-STYLE-left_mode (car borders)))
                  (list "color" (cons "rgb" (BORDER-STYLE-left_color (car borders))))))
                '(("left")))
            (if (BORDER-STYLE-right_mode (car borders))
                (list
                 (list
                  "right"
                  (cons "style" (BORDER-STYLE-right_mode (car borders)))
                  (list "color" (cons "rgb" (BORDER-STYLE-right_color (car borders))))))
                '(("right")))
            (if (BORDER-STYLE-top_mode (car borders))
                (list
                 (list
                  "top"
                  (cons "style" (BORDER-STYLE-top_mode (car borders)))
                  (list "color" (cons "rgb" (BORDER-STYLE-top_color (car borders))))))
                '(("top")))
            (if (BORDER-STYLE-bottom_mode (car borders))
                (list
                 (list
                  "bottom"
                  (cons "style" (BORDER-STYLE-bottom_mode (car borders)))
                  (list "color" (cons "rgb" (BORDER-STYLE-bottom_color (car borders))))))
                '(("bottom")))
            '(("diagonal")))
           result_list))
         (reverse result_list)))))

(define (from-borders xml_hash)
  (let loop ([loop_count 0]
             [border_list '()])
    (if (< loop_count (hash-ref xml_hash "styleSheet1.borders1.border's count" 0))
        (let* ([prefix (format "styleSheet1.borders1.border~a" (add1 loop_count))]
               [border_style
                (BORDER-STYLE
                 (let ([left_color (hash-ref xml_hash (format "~a.left1.color1.rgb1" prefix) "")])
                   (if (string=? left_color "") #f left_color))
                 (let ([left_style (hash-ref xml_hash (format "~a.left1.style1" prefix) "")])
                   (if (string=? left_style "") #f left_style))
                 (let ([right_color (hash-ref xml_hash (format "~a.right1.color1.rgb1" prefix) "")])
                   (if (string=? right_color "") #f right_color))
                 (let ([right_style (hash-ref xml_hash (format "~a.right1.style1" prefix) "")])
                   (if (string=? right_style "") #f right_style))
                 (let ([top_color (hash-ref xml_hash (format "~a.top1.color1.rgb1" prefix) "")])
                   (if (string=? top_color "") #f top_color))
                 (let ([top_style (hash-ref xml_hash (format "~a.top1.style1" prefix) "")])
                   (if (string=? top_style "") #f top_style))
                 (let ([bottom_color (hash-ref xml_hash (format "~a.bottom1.color1.rgb1" prefix) "")])
                   (if (string=? bottom_color "") #f bottom_color))
                 (let ([bottom_style (hash-ref xml_hash (format "~a.bottom1.style1" prefix) "")])
                   (if (string=? bottom_style "") #f bottom_style))
                 )])
          (loop (add1 loop_count) (cons border_style border_list)))
        (set-STYLES-border_list! (*STYLES*) (reverse border_list)))))

