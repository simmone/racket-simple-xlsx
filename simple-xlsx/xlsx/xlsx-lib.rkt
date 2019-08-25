#lang racket

(require "../lib/lib.rkt")

(provide (contract-out
          [check-data-list (-> any/c boolean?)]
          ))

(define (check-data-list data_list)
  (when (not (list? data_list))
        (error "data is not list type"))
  
  (when (equal? data_list '())
        (error "data has no children list"))
  
  (let loop ([loop_list data_list]
             [child_length -1])
    (when (not (null? loop_list))
          (when (not (list? (car loop_list)))
                (error "data's children is not list type"))
          
          (when (and
                 (not (= child_length -1))
                 (not (= child_length (length (car loop_list)))))
                (error "data's children's length is not consistent."))
          
          (loop (cdr loop_list) (length (car loop_list)))))
  #t)
