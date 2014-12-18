#lang racket

(require rackunit/text-ui)

(require "lib/lib-test.rkt")

(require "test1/test1-test.rkt")

(require "test2/test2-test.rkt")

(require "test3/test3-test.rkt")

(require "test4/test4-test.rkt")

(require "test5/test5-test.rkt")

(require "test6/test6-test.rkt")

(run-tests test-lib)

(run-tests test-test1)

(run-tests test-test2)

(run-tests test-test3)

(run-tests test-test4)

(run-tests test-test5)

(run-tests test-test6)

