#lang racket

(require simple-xml)

(require racket/date)

(require rackunit/text-ui rackunit)

(require "../../../../../xlsx/xlsx.rkt")
(require "../../../../../sheet/sheet.rkt")
(require "../../../../../style/style.rkt")
(require "../../../../../style/set-styles.rkt")
(require "../../../../../lib/lib.rkt")
(require "../../../../../lib/sheet-lib.rkt")

(require"../../../../../xl/worksheets/worksheet.rkt")

(require racket/runtime-path)
(define-runtime-path rows1_file "rows1.xml")
(define-runtime-path rows2_file "rows2.xml")

(define test-rows
  (test-suite
   "test-rows"

   (test-case
    "test-to-rows1"

    (with-xlsx
     (lambda ()
       (add-data-sheet "Sheet1"
                       (list
                        (list "month1" "month2" "month3" "month1" "real")
                        (list 201601 100 110 1110 6.9)
                        (list
                         (seconds->date (find-seconds 0 0 0 17 9 2018 #f))
                         (seconds->date (find-seconds 0 0 0 17 9 2018 #f))
                         (seconds->date (find-seconds 0 0 0 17 9 2018 #f))
                         (seconds->date (find-seconds 0 0 0 17 9 2018 #f))
                         (seconds->date (find-seconds 0 0 0 17 9 2018 #f)))
                        ))

       (squash-shared-strings-map)

       (with-sheet
        (lambda ()
          (set-row-range-height "2-3" 5)

          (call-with-input-file rows1_file
            (lambda (expected)
              (call-with-input-string
               (lists->xml_content (to-rows))
               (lambda (actual)
                 (check-lines? expected actual)))))
          )))))

   (test-case
    "test-from-rows1"

    (with-xlsx
     (lambda ()
       (add-data-sheet "Sheet1" '(("none")))

       (hash-set! (XLSX-shared_string->index_map (*XLSX*)) "month1" 0)
       (hash-set! (XLSX-shared_index->string_map (*XLSX*)) 0 "month1")

       (hash-set! (XLSX-shared_string->index_map (*XLSX*)) "month2" 1)
       (hash-set! (XLSX-shared_index->string_map (*XLSX*)) 1 "month2")

       (hash-set! (XLSX-shared_string->index_map (*XLSX*)) "month3" 2)
       (hash-set! (XLSX-shared_index->string_map (*XLSX*)) 2 "month3")

       (hash-set! (XLSX-shared_string->index_map (*XLSX*)) "real" 3)
       (hash-set! (XLSX-shared_index->string_map (*XLSX*)) 3 "real")

       (with-sheet
        (lambda ()
          (let ([xml_hash
                 (xml->hash (open-input-string (format "<worksheet>~a~a</worksheet>"
                                                       "<dimension ref=\"A1:E3\"/>"
                                                       (file->string rows1_file)
                                                       )))])
            (from-work-sheet-head xml_hash)

            (from-rows xml_hash)

            (check-equal? (hash-count (DATA-SHEET-cell->value_hash (*CURRENT_SHEET*))) 15)

            (check-equal? (hash-ref (DATA-SHEET-cell->value_hash (*CURRENT_SHEET*)) "A1") "month1")
            (check-equal? (hash-ref (DATA-SHEET-cell->value_hash (*CURRENT_SHEET*)) "B1") "month2")
            (check-equal? (hash-ref (DATA-SHEET-cell->value_hash (*CURRENT_SHEET*)) "C1") "month3")
            (check-equal? (hash-ref (DATA-SHEET-cell->value_hash (*CURRENT_SHEET*)) "D1") "month1")
            (check-equal? (hash-ref (DATA-SHEET-cell->value_hash (*CURRENT_SHEET*)) "E1") "real")

            (check-equal? (hash-ref (DATA-SHEET-cell->value_hash (*CURRENT_SHEET*)) "A2") 201601)
            (check-equal? (hash-ref (DATA-SHEET-cell->value_hash (*CURRENT_SHEET*)) "B2") 100)
            (check-equal? (hash-ref (DATA-SHEET-cell->value_hash (*CURRENT_SHEET*)) "C2") 110)
            (check-equal? (hash-ref (DATA-SHEET-cell->value_hash (*CURRENT_SHEET*)) "D2") 1110)
            (check-equal? (hash-ref (DATA-SHEET-cell->value_hash (*CURRENT_SHEET*)) "E2") 6.9)

            (check-equal? (hash-ref (DATA-SHEET-cell->value_hash (*CURRENT_SHEET*)) "A3") 43360)
            (check-equal? (hash-ref (DATA-SHEET-cell->value_hash (*CURRENT_SHEET*)) "B3") 43360)
            (check-equal? (hash-ref (DATA-SHEET-cell->value_hash (*CURRENT_SHEET*)) "C3") 43360)
            (check-equal? (hash-ref (DATA-SHEET-cell->value_hash (*CURRENT_SHEET*)) "D3") 43360)
            (check-equal? (hash-ref (DATA-SHEET-cell->value_hash (*CURRENT_SHEET*)) "E3") 43360)

            ))))))

   (test-case
    "test-to-rows2"

    (with-xlsx
     (lambda ()
       (add-data-sheet "Sheet1"
                       (list
                        (list "month1" "month2" "month3" "month1" "real")
                        (list 201601 100 110 1110 6.9)
                        (list
                         (seconds->date (find-seconds 0 0 0 17 9 2018 #f))
                         (seconds->date (find-seconds 0 0 0 17 9 2018 #f))
                         (seconds->date (find-seconds 0 0 0 17 9 2018 #f))
                         (seconds->date (find-seconds 0 0 0 17 9 2018 #f))
                         (seconds->date (find-seconds 0 0 0 17 9 2018 #f)))
                        )
                       "C3")

       (squash-shared-strings-map)

       (with-sheet
        (lambda ()
          (set-row-range-height "4-5" 5)

          (call-with-input-file rows2_file
            (lambda (expected)
              (call-with-input-string
               (lists->xml_content (to-rows))
               (lambda (actual)
                 (check-lines? expected actual)))))
          )))))

   (test-case
    "test-from-rows2"

    (with-xlsx
     (lambda ()
       (add-data-sheet "Sheet1" '(("none")))

       (hash-set! (XLSX-shared_string->index_map (*XLSX*)) "month1" 0)
       (hash-set! (XLSX-shared_index->string_map (*XLSX*)) 0 "month1")

       (hash-set! (XLSX-shared_string->index_map (*XLSX*)) "month2" 1)
       (hash-set! (XLSX-shared_index->string_map (*XLSX*)) 1 "month2")

       (hash-set! (XLSX-shared_string->index_map (*XLSX*)) "month3" 2)
       (hash-set! (XLSX-shared_index->string_map (*XLSX*)) 2 "month3")

       (hash-set! (XLSX-shared_string->index_map (*XLSX*)) "real" 3)
       (hash-set! (XLSX-shared_index->string_map (*XLSX*)) 3 "real")

       (with-sheet
        (lambda ()
          (let ([xml_hash
                 (xml->hash (open-input-string (format "<worksheet>~a~a</worksheet>"
                                                       "<dimension ref=\"C3:G5\"/>"
                                                       (file->string rows2_file)
                                                       )))])

            (hash-clear! (DATA-SHEET-cell->value_hash (*CURRENT_SHEET*)))

            (from-work-sheet-head xml_hash)

            (from-rows xml_hash)

            (check-equal? (hash-count (DATA-SHEET-cell->value_hash (*CURRENT_SHEET*))) 15)

            (check-equal? (hash-ref (DATA-SHEET-cell->value_hash (*CURRENT_SHEET*)) "C3") "month1")
            (check-equal? (hash-ref (DATA-SHEET-cell->value_hash (*CURRENT_SHEET*)) "D3") "month2")
            (check-equal? (hash-ref (DATA-SHEET-cell->value_hash (*CURRENT_SHEET*)) "E3") "month3")
            (check-equal? (hash-ref (DATA-SHEET-cell->value_hash (*CURRENT_SHEET*)) "F3") "month1")
            (check-equal? (hash-ref (DATA-SHEET-cell->value_hash (*CURRENT_SHEET*)) "G3") "real")

            (check-equal? (hash-ref (DATA-SHEET-cell->value_hash (*CURRENT_SHEET*)) "C4") 201601)
            (check-equal? (hash-ref (DATA-SHEET-cell->value_hash (*CURRENT_SHEET*)) "D4") 100)
            (check-equal? (hash-ref (DATA-SHEET-cell->value_hash (*CURRENT_SHEET*)) "E4") 110)
            (check-equal? (hash-ref (DATA-SHEET-cell->value_hash (*CURRENT_SHEET*)) "F4") 1110)
            (check-equal? (hash-ref (DATA-SHEET-cell->value_hash (*CURRENT_SHEET*)) "G4") 6.9)

            (check-equal? (hash-ref (DATA-SHEET-cell->value_hash (*CURRENT_SHEET*)) "C5") 43360)
            (check-equal? (hash-ref (DATA-SHEET-cell->value_hash (*CURRENT_SHEET*)) "D5") 43360)
            (check-equal? (hash-ref (DATA-SHEET-cell->value_hash (*CURRENT_SHEET*)) "E5") 43360)
            (check-equal? (hash-ref (DATA-SHEET-cell->value_hash (*CURRENT_SHEET*)) "F5") 43360)
            (check-equal? (hash-ref (DATA-SHEET-cell->value_hash (*CURRENT_SHEET*)) "G5") 43360)

            ))))))
   ))

(run-tests test-rows)
