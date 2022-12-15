#lang racket

(require simple-xml)

(require rackunit/text-ui rackunit)

(require "../../../../../xlsx/xlsx.rkt")
(require "../../../../../sheet/sheet.rkt")
(require "../../../../../style/style.rkt")
(require "../../../../../style/set-styles.rkt")
(require "../../../../../lib/lib.rkt")

(require"../../../../../xl/worksheets/worksheet.rkt")

(require racket/runtime-path)
(define-runtime-path sheet_view_file "sheet_view.xml")
(define-runtime-path sheet_view_other_file "sheet_view_other.xml")
(define-runtime-path sheet_view_freeze_1_1_file "sheet_view_freeze_1_1.xml")
(define-runtime-path sheet_view_freeze_1_0_file "sheet_view_freeze_1_0.xml")
(define-runtime-path sheet_view_freeze_0_1_file "sheet_view_freeze_0_1.xml")
(define-runtime-path sheet_view_freeze_3_2_file "sheet_view_freeze_3_2.xml")

(define test-worksheet
  (test-suite
   "test-worksheet"

   (test-case
    "test-to-sheet-view"

    (with-xlsx
     (lambda ()
       (add-data-sheet "Sheet1"
                       '(("month1" "month2" "month3" "month4" "real") (201601 100 110 1110 6.9)))
       (add-data-sheet "Sheet2" '((1)))
       (add-data-sheet "Sheet3" '((1)))
       (add-data-sheet "Sheet4" '((1)))
       (add-chart-sheet "Chart1" 'LINE "Chart1" '())
       (add-chart-sheet "Chart2" 'LINE "Chart2" '())
       (add-chart-sheet "Chart3" 'LINE "Chart3" '())

       (with-sheet-ref
        0
        (lambda ()
          (call-with-input-file sheet_view_file
            (lambda (expected)
              (call-with-input-string
               (lists->xml_content (to-sheet-view))
               (lambda (actual)
                 (check-lines? expected actual)))))))

       (with-sheet-ref
        1
        (lambda ()
          (call-with-input-file sheet_view_other_file
            (lambda (expected)
              (call-with-input-string
               (lists->xml_content (to-sheet-view))
               (lambda (actual)
                 (check-lines? expected actual)))))))

       (with-sheet-ref
        0
        (lambda ()
          (set-freeze-row-col-range 1 0)

          (call-with-input-file sheet_view_freeze_1_0_file
            (lambda (expected)
              (call-with-input-string
               (lists->xml_content (to-sheet-view))
               (lambda (actual)
                 (check-lines? expected actual)))))))

       (with-sheet-ref
        1
        (lambda ()
          (set-freeze-row-col-range 0 1)

          (call-with-input-file sheet_view_freeze_0_1_file
            (lambda (expected)
              (call-with-input-string
               (lists->xml_content (to-sheet-view))
               (lambda (actual)
                 (check-lines? expected actual)))))))

       (with-sheet-ref
        2
        (lambda ()
          (set-freeze-row-col-range 1 1)

          (call-with-input-file sheet_view_freeze_1_1_file
            (lambda (expected)
              (call-with-input-string
               (lists->xml_content (to-sheet-view))
               (lambda (actual)
                 (check-lines? expected actual)))))))

       (with-sheet-ref
        3
        (lambda ()
          (set-freeze-row-col-range 3 2)

          (call-with-input-file sheet_view_freeze_3_2_file
            (lambda (expected)
              (call-with-input-string
               (lists->xml_content (to-sheet-view))
               (lambda (actual)
                 (check-lines? expected actual)))))))
       )))

   (test-case
    "test-from-sheet-view"

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
          (from-sheet-view
           (xml->hash (open-input-string (format "<worksheet><sheetViews>~a</sheetViews></worksheet>" (file->string sheet_view_file)))))

          (check-equal? (SHEET-STYLE-freeze_range (*CURRENT_SHEET_STYLE*)) '(0 . 0))))

       (with-sheet
        (lambda ()
          (from-sheet-view
           (xml->hash (open-input-string (format "<worksheet><sheetViews>~a</sheetViews></worksheet>" (file->string sheet_view_other_file)))))

          (check-equal? (SHEET-STYLE-freeze_range (*CURRENT_SHEET_STYLE*)) '(0 . 0))))

       (with-sheet
        (lambda ()
          (from-sheet-view
           (xml->hash (open-input-string (format "<worksheet><sheetViews>~a</sheetViews></worksheet>" (file->string sheet_view_freeze_1_0_file)))))

          (check-equal? (SHEET-STYLE-freeze_range (*CURRENT_SHEET_STYLE*)) '(1 . 0))))

       (with-sheet
        (lambda ()
          (from-sheet-view
           (xml->hash (open-input-string (format "<worksheet><sheetViews>~a</sheetViews></worksheet>" (file->string sheet_view_freeze_1_1_file)))))

          (check-equal? (SHEET-STYLE-freeze_range (*CURRENT_SHEET_STYLE*)) '(1 . 1))))

       (with-sheet
        (lambda ()
          (from-sheet-view
           (xml->hash (open-input-string (format "<worksheet><sheetViews>~a</sheetViews></worksheet>" (file->string sheet_view_freeze_0_1_file)))))

          (check-equal? (SHEET-STYLE-freeze_range (*CURRENT_SHEET_STYLE*)) '(0 . 1))))

       (with-sheet
        (lambda ()
          (from-sheet-view
           (xml->hash (open-input-string (format "<worksheet><sheetViews>~a</sheetViews></worksheet>" (file->string sheet_view_freeze_3_2_file)))))

          (check-equal? (SHEET-STYLE-freeze_range (*CURRENT_SHEET_STYLE*)) '(3 . 2))))

       )))

   ))

(run-tests test-worksheet)
