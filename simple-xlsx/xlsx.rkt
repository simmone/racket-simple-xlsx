#lang racket

(provide (contract-out
          [xlsx% class?]
          [struct sheet
                  (
                  (name string?)
                  (seq exact-nonnegative-integer?)
                  (type symbol?)
                  (typeSeq exact-nonnegative-integer?)
                  (data list?)
                  )]
          ))

(struct sheet ([name #:mutable] [seq #:mutable] [type #:mutable] [typeSeq #:mutable] [data #:mutable]))

(struct sheetData ([data_list #:mutable] [col_attr_hash #:mutable]))

(struct lineChartData ([topic #:mutable] [x_data #:mutable] [y_data_hash #:mutable]))

(struct colAttr ([width #:mutable] [back_color #:mutable]))

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
                              sheet_data)))))
         
         (define/public (sheet-ref sheet_seq)
           (list-ref sheets sheet_seq))

         ))

