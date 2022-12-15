#lang racket

(provide (contract-out
          [squeeze-range-hash (-> (hash/c natural? number?) (listof (list/c natural? natural? number?)))]
          [rgb? (-> string? boolean?)]
          ))

(define (rgb? color_string)
  (if (or
       (regexp-match #px"^([0-9]|[a-zA-Z]){6}$" color_string)
       (regexp-match #px"^([0-9]|[a-zA-Z]){8}$" color_string))
      #t
      #f))

(define (squeeze-range-hash range_hash)
  (let loop ([range_list (sort #:key car (hash->list range_hash) <)]
             [loop_start_index -1]
             [loop_end_index -1]
             [loop_val -1]
             [result_list '()])
    (if (not (null? range_list))
        (let* ([range (car range_list)]
               [index (car range)]
               [val (cdr range)])
          (if (and (= index (add1 loop_end_index)) (= val loop_val))
              (loop
               (cdr range_list)
               loop_start_index
               (add1 loop_end_index)
               val
               result_list)
              (if (= loop_start_index -1)
                  (loop (cdr range_list) index index val result_list)
                  (loop (cdr range_list) index index val (cons (list loop_start_index loop_end_index loop_val) result_list)))))
        (if (= loop_start_index -1)
            '()
            (reverse
             (cons (list loop_start_index loop_end_index loop_val) result_list))))))
