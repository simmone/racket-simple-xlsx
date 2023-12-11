#lang info

(define scribblings
  '(("scribble/simple-xlsx.scrbl" (multi-page) (tool 100))))

(define compile-omit-paths '("tests" "verify"))
(define test-omit-paths '("lib" "*.rkt" "*.scrbl" "lib" "example" "reader" "writer" "xlsx" "sheet" "verify"))


