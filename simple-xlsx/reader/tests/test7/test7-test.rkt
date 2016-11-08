#lang racket

(require rackunit/text-ui)

(require rackunit "../../../main.rkt")

(require racket/runtime-path)
(define-runtime-path test_file "test7.xlsx")

(define test-test7
  (test-suite
   "test-test7"

   (with-input-from-xlsx-file
    test_file
    (lambda (xlsx)
      (test-case
       "test-get-sheet-data"

       (load-sheet (car (get-sheet-names xlsx)) xlsx)
       (check-equal? (get-cell-value "A2" xlsx) "er")
       (check-equal? (get-cell-value "A1" xlsx) "1、请按在行离行标识分成两张表\r\n2、每张表请按设备类型、日均交易笔数排序")
      )))))

(run-tests test-test7)
