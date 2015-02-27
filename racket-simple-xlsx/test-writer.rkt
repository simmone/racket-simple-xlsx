#lang racket

(require "writer.rkt")

(let ([data_list '(((1 2 3) (4 5 6)) (("1" "2" "4") ("5")) (("1")))])
  (write-xlsx-file data_list '("员工" "经理") "haha"))
