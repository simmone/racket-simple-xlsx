#lang racket

(require rackunit/text-ui)

(require rackunit "../../main.rkt")

(let ([xlsx (new xlsx-data%)]
      [rows '()])

  (printf "t01\n")
  (with-input-from-file "test.data"
    (lambda ()
      (let loop ([line (read-line)]
                 [count 1])
        (when (not (eof-object? line))
              (when (= (remainder count 10000) 0)
                    (printf "~a\n" count))
              (set! rows `(,@rows ,(list (list (regexp-replace* #rx"\n|\r" line "")))))
              (loop (read-line) (add1 count))))))

    (printf "t02\n")
    (send xlsx add-sheet rows "Sheet1")

    (printf "t03\n")
  (write-xlsx-file xlsx "test1.xlsx"))
