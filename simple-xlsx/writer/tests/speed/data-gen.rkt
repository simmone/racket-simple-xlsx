#lang racket

(with-output-to-file "test.data"
  #:exists 'replace
  (lambda ()
    (let loop ([count 1])
      (when (<= count 100000)
            (printf "~a\n" count)
            (loop (add1 count))))))
      