#lang racket

(require "writer.rkt")

(let ([data_list '((("chenxiao" "陈晓")) ((1 2 3 4)) ())])
  (write-xlsx-file data_list #f "/Users/simmone/Downloads/haha.xlsx"))
