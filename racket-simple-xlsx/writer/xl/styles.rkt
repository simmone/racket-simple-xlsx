#lang at-exp racket/base

(require racket/port)
(require racket/list)
(require racket/contract)

;; strings list convert to (string . place) hash
(provide (contract-out
          [write-styles (-> string?)]
          ))

(define S string-append)

(define (write-styles) @S{
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fonts count="1"><font /></fonts><fills count="1"><fill /></fills><borders count="1"><border /></borders><cellStyleXfs count="1"><xf /></cellStyleXfs><cellXfs count="2"><xf /><xf fontId="1" /></cellXfs></styleSheet>
})
