#lang info

(define version "3.1")

(define license
  '(Apache-2.0 OR MIT))

(define collection 'multi)

(define deps '("base"
               "rackunit-lib"
               "fast-xml"
               "scribble-lib"
               ))

(define test-omit-paths '("info.rkt"))
