#lang info

(define scribblings
  '(("simple-xlsx.scrbl" (multi-page) (tool 100))))

(define compile-omit-paths '("tests"))
(define test-omit-paths '("lib" "*.rkt" "*.scrbl" "lib" "example" "reader" "writer" "xlsx" "sheet"))


