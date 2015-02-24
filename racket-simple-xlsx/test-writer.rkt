#lang racket

(require "writer.rkt")

(let ([data_list '(((1 2 3) (4 5 6)) (("1" "2" "4") ("5")))])
  (write-xlsx-file data_list "haha"))
