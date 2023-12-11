#lang racket

(require simple-xml)

(require rackunit/text-ui rackunit)

(require "../../../../../xlsx/xlsx.rkt")
(require "../../../../../sheet/sheet.rkt")
(require "../../../../../style/style.rkt")
(require "../../../../../style/styles.rkt")
(require "../../../../../style/set-styles.rkt")
(require "../../../../../lib/lib.rkt")

(require"../../../../../xl/worksheets/worksheet.rkt")

(require racket/runtime-path)
(define-runtime-path sheet_cols1_file "sheet_cols1.xml")
(define-runtime-path sheet_cols2_file "sheet_cols2.xml")
(define-runtime-path sheet_no_cols_file "sheet_no_cols.xml")

(define test-worksheet
  (test-suite
   "test-worksheet"

   (test-case
    "test-to-col-width-style"

    (with-xlsx
     (lambda ()
       (add-data-sheet "Sheet1"
                       '(("month1" "month2" "month3" "month4" "real") (201601 100 110 1110 6.9)))
       (add-data-sheet "Sheet2" '((1)))
       (add-data-sheet "Sheet3" '((1)))
       (add-chart-sheet "Chart1" 'LINE "Chart1" '())
       (add-chart-sheet "Chart2" 'LINE "Chart2" '())
       (add-chart-sheet "Chart3" 'LINE "Chart3" '())

       (with-sheet
        (lambda ()
          (set-col-range-width "1-3" 5)
          (set-col-range-width "C-F" 6)
          (set-col-range-width "F-10" 7)

          (call-with-input-file sheet_cols1_file
            (lambda (expected)
              (call-with-input-string
               (lists->xml_content (to-cols))
               (lambda (actual)
                 (check-lines? expected actual)))))))

       (with-sheet-ref
        1
        (lambda ()
          (set-col-range-width "F-10" 7)
          
          (call-with-input-file sheet_cols2_file
            (lambda (expected)
              (call-with-input-string
               (lists->xml_content (to-cols))
               (lambda (actual)
                 (check-lines? expected actual)))))))

       (with-sheet-ref
        2
        (lambda ()
          (call-with-input-file sheet_no_cols_file
            (lambda (expected)
              (call-with-input-string
               (lists->xml_content (to-cols))
               (lambda (actual)
                 (check-lines? expected actual)))))))
          )))

   (test-case
    "test-from-col-width-style"

    (with-xlsx
     (lambda ()
       (add-data-sheet "Sheet1" '((1)))
       (add-data-sheet "Sheet2" '((1)))
       (add-data-sheet "Sheet3" '((1)))
       (add-data-sheet "Sheet4" '((1)))
       (add-chart-sheet "Chart1" 'LINE "Chart1" '())
       (add-chart-sheet "Chart2" 'LINE "Chart2" '())
       (add-chart-sheet "Chart3" 'LINE "Chart3" '())

       (with-sheet
        (lambda ()
          (from-col-width-style
           (xml->hash (open-input-string (format "<worksheet>~a</worksheet>" (file->string sheet_cols1_file)))))

          (check-equal? (hash-count (SHEET-STYLE-col->width_map (*CURRENT_SHEET_STYLE*))) 10)
          (check-equal? (hash-ref (SHEET-STYLE-col->width_map (*CURRENT_SHEET_STYLE*)) 1) 5)
          (check-equal? (hash-ref (SHEET-STYLE-col->width_map (*CURRENT_SHEET_STYLE*)) 2) 5)
          (check-equal? (hash-ref (SHEET-STYLE-col->width_map (*CURRENT_SHEET_STYLE*)) 3) 6)
          (check-equal? (hash-ref (SHEET-STYLE-col->width_map (*CURRENT_SHEET_STYLE*)) 4) 6)
          (check-equal? (hash-ref (SHEET-STYLE-col->width_map (*CURRENT_SHEET_STYLE*)) 5) 6)
          (check-equal? (hash-ref (SHEET-STYLE-col->width_map (*CURRENT_SHEET_STYLE*)) 6) 7)
          (check-equal? (hash-ref (SHEET-STYLE-col->width_map (*CURRENT_SHEET_STYLE*)) 7) 7)
          (check-equal? (hash-ref (SHEET-STYLE-col->width_map (*CURRENT_SHEET_STYLE*)) 8) 7)
          (check-equal? (hash-ref (SHEET-STYLE-col->width_map (*CURRENT_SHEET_STYLE*)) 9) 7)
          (check-equal? (hash-ref (SHEET-STYLE-col->width_map (*CURRENT_SHEET_STYLE*)) 10) 7)

       (with-sheet-ref
        1
        (lambda ()
          (from-col-width-style
           (xml->hash (open-input-string (format "<worksheet>~a</worksheet>" (file->string sheet_cols2_file)))))

          (check-equal? (hash-count (SHEET-STYLE-col->width_map (*CURRENT_SHEET_STYLE*))) 5)
          (check-equal? (hash-ref (SHEET-STYLE-col->width_map (*CURRENT_SHEET_STYLE*)) 6) 7)
          (check-equal? (hash-ref (SHEET-STYLE-col->width_map (*CURRENT_SHEET_STYLE*)) 7) 7)
          (check-equal? (hash-ref (SHEET-STYLE-col->width_map (*CURRENT_SHEET_STYLE*)) 8) 7)
          (check-equal? (hash-ref (SHEET-STYLE-col->width_map (*CURRENT_SHEET_STYLE*)) 9) 7)
          (check-equal? (hash-ref (SHEET-STYLE-col->width_map (*CURRENT_SHEET_STYLE*)) 10) 7)
       )))))))

   ))

(run-tests test-worksheet)
