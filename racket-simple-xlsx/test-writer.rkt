#lang racket

(require "writer.rkt")

(let ([data_list '((("chenxiao")) () ())])
  (write-xlsx-file data_list #f "/Users/simmone/Downloads/haha.xlsx"))
