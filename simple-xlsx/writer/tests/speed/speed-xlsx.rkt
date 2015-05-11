#lang racket

(require rackunit/text-ui)

(require rackunit "../../../main.rkt")

(let ([xlsx (new xlsx-data%)])
  (with-input-from-file "test.data"
    (lambda ()
      (let loop ([rows '()]
                 [line (read-line)]
                 [count 1])
                 
        (if (not (eof-object? line))
            (begin
              (when (= (remainder count 10000) 0)
                    (printf "~a\n" count))
              (loop (cons (list (regexp-replace* #rx"\n|\r" line "")) rows) (read-line) (add1 count)))
            (send xlsx add-sheet rows "Sheet1")))))

  (write-xlsx-file xlsx "test1.xlsx"))
