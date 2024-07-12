#lang racket

(require fast-xml
         rackunit/text-ui
         rackunit
         "../../../../../xlsx/xlsx.rkt"
         "../../../../../sheet/sheet.rkt"
         "../../../../../lib/lib.rkt"
         "../../../../../xl/worksheets/worksheet.rkt"
         racket/runtime-path)

(define-runtime-path work_sheet_tail_phoneticPr_file "work_sheet_tail_phoneticPr.xml")
(define-runtime-path work_sheet_tail_pageMargins_file "work_sheet_tail_pageMargins.xml")
(define-runtime-path work_sheet_tail_pageSetup_file "work_sheet_tail_pageSetup.xml")

(define test-worksheet
  (test-suite
   "test-worksheet"

   (test-case
    "test-work-sheet-tail"

    (with-xlsx
     (lambda ()
       (add-data-sheet "Sheet1"
                       '(("month1" "month2" "month3" "month4" "real") (201601 100 110 1110 6.9)))

       (parameterize
           ([*CURRENT_SHEET* (car (XLSX-sheet_list (*XLSX*)))])

         (call-with-input-file work_sheet_tail_phoneticPr_file
           (lambda (expected)
             (call-with-input-string
              (lists-to-xml_content (first (work-sheet-tail)))
              (lambda (actual)
                (check-lines? expected actual)))))

         (call-with-input-file work_sheet_tail_pageMargins_file
           (lambda (expected)
             (call-with-input-string
              (lists-to-xml_content (second (work-sheet-tail)))
              (lambda (actual)
                (check-lines? expected actual)))))

         (call-with-input-file work_sheet_tail_pageSetup_file
           (lambda (expected)
             (call-with-input-string
              (lists-to-xml_content (third (work-sheet-tail)))
              (lambda (actual)
                (check-lines? expected actual)))))
         ))))
   ))

(run-tests test-worksheet)
