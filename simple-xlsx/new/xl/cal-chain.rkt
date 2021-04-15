#lang at-exp racket/base

(require racket/port)
(require racket/list)
(require racket/contract)

(provide (contract-out
          [write-cal-chain (-> string?)]
          ))

(define S string-append)
 
(define (write-cal-chain) @S{
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<calcChain xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><c r="C2" i="1"/></calcChain>
})
