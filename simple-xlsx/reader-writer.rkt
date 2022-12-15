#lang racket

(require racket/date)

(provide (contract-out
          [read-and-write-xlsx (-> path-string? path-string? procedure? any)]
          ))

(require "xlsx/xlsx.rkt")
(require "lib/lib.rkt")
(require "reader.rkt")
(require "writer.rkt")

(define (read-and-write-xlsx from_file_name to_file_name user_proc)
  (when (not (file-exists? from_file_name))
    (error (format "xlsx file [~a] not exists!" from_file_name)))

  (with-xlsx
   (lambda ()
     (dynamic-wind
         (lambda () (set-XLSX-xlsx_dir! (*XLSX*) (make-temporary-file "xlsx_tmp_~a" 'directory ".")))
         (lambda ()
           (unzip-xlsx from_file_name (XLSX-xlsx_dir (*XLSX*)))

           (read-batch)

           (user_proc)

           (when (file-exists? to_file_name)
             (delete-file to_file_name))

           (write-batch)

           (zip-xlsx to_file_name (XLSX-xlsx_dir (*XLSX*))))
         (lambda ()
           (delete-directory/files (XLSX-xlsx_dir (*XLSX*)))
           )
         ))))
