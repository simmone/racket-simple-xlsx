#lang racket

(require rackunit/text-ui)

(with-output-to-file "test1.csv"
  #:exists 'replace
  (lambda ()
    (with-input-from-file "test.data"
      (lambda ()
        (let loop ([line (read-line)]
                   [count 1])
          (when (not (eof-object? line))
                (printf "~a\n" (regexp-replace* #rx"\n|\r" line ""))
                (loop (read-line) (add1 count))))))))
