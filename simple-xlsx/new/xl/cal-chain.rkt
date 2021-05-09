#lang racket

(require simple-xml)

(provide (contract-out
          [cal-chain (-> string?)]
          ))

(define (cal-chain) {
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<calcChain xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><c r="C2" i="1"/></calcChain>
})
