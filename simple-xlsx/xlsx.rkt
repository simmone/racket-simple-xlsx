#lang racket

(provide (contract-out
          [xlsx% class?]
          [struct sheet
                  (
                  (name string?)
                  (seq exact-nonnegative-integer?)
                  (type symbol?)
                  (typeSeq exact-nonnegative-integer?)
                  (content (or/c data-sheet? line-chart-sheet?))
                  )]
          [struct data-sheet
                  (
                   (rows list?)
                   (width_hash hash?)
                   (color_hash hash?)
                   )]
          [struct line-chart-sheet
                  (
                   (topic string?)
                   (x_data list?)
                   (y_data_list list?)
                   )]
          ))

(struct sheet ([name #:mutable] [seq #:mutable] [type #:mutable] [typeSeq #:mutable] [content #:mutable]))

(struct data-sheet ([rows #:mutable] [width_hash #:mutable] [color_hash #:mutable]))
(struct colAttr ([width #:mutable] [back_color #:mutable]))

(struct line-chart-sheet ([topic #:mutable] [x_data #:mutable] [y_data_list #:mutable]))
(struct data-serial ([topic #:mutable] [data_list #:mutable]))

(define xlsx%
  (class object%
         (super-new)
         
         (field 
          [sheets '()]
          )
         
         (define/public (add-data-sheet sheet_name sheet_data)
           (let* ([sheet_length (length sheets)]
                  [seq (add1 sheet_length)]
                  [type_seq (add1 (length (filter (lambda (rec) (eq? (sheet-type rec) 'data)) sheets)))])
             (set! sheets `(,@sheets
                            ,(sheet
                              sheet_name
                              seq
                              'data
                              type_seq
                              (data-sheet sheet_data (make-hash) (make-hash)))))))
         
         (define/public (sheet-ref sheet_seq)
           (list-ref sheets sheet_seq))
         
         (define/public (set-sheet-col-width! sheet col_range width)
           (hash-set! (data-sheet-width_hash (sheet-content sheet)) col_range width))

         (define/public (set-sheet-col-color! sheet col_range color)
           (hash-set! (data-sheet-color_hash (sheet-content sheet)) col_range color))
         ))

