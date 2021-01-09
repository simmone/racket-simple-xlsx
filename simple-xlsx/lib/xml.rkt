#lang racket

(require xml)
(require xml/xexpr)

(require detail)

(provide (contract-out
          [load-xml-hash (-> path-string? (listof symbol?) hash?)]
          ))

(define (load-xml-hash xml sym_list)
  (detail
   #:formats? '("xml_decode.pdf")
   (lambda ()
     (detail-page
      #:line_break_length? 110
      #:font_size? 'small
      (lambda ()
        (detail-h1 "Decode XML Process")
        (detail-line (format "XML File:~a" xml))

        (call-with-input-file
            xml
          (lambda (origin_port)
            ;; remove all the spaces
            (call-with-input-string
             (regexp-replace* #rx"> *<"
                              (regexp-replace* #rx"\n|\r" (port->string origin_port) "")
                              "><")
             (lambda (filtered_port)
               (let ([xml_hash (make-hash)]
                     [sym_hash (make-hash)]
                     )

                 ;; which syms need be count and convert to list, turn them to map to convenient to check
                 (map (lambda (sym) (hash-set! sym_hash sym #f)) sym_list)
                 (detail-line (format "List Symbol:~a" sym_hash))

                 ;; load xml file to xexpr to start parsing
                 ;; ancester_prefix means parent node name, start from #f
                 (let loop-node ([ancester_prefix #f]
                                 [xml_list (list (xml->xexpr (document-element (read-xml filtered_port))))])

                   (detail-line "")
                   (detail-line "***************************************")
                   (detail-line (format "ancester_prefix:[~a]\n" ancester_prefix))
                   (detail-line (format "xml_list:[~a]\n" xml_list))
                   (detail-line "***************************************")
                   
                   (if (not (null? xml_list))
                       (let ([node (car xml_list)])
                         (detail-line (format "node:[~a]" node))
                         (if (not (list? node))
                             (if (null? node)
                                 (begin
                                   (detail-line "null node, next loop")
                                   (loop-node node (cdr xml_list)))
                                 (begin
                                   (detail-line "node is a symbol")
                                   (detail-line (format "is_list?:[~a]" (hash-has-key? sym_hash ancester_prefix)))
                                   (if (hash-has-key? sym_hash ancester_prefix) ;; prefix is the sym we expect it to be a list, count it
                                       (let ([count_sym (format "~a.count" ancester_prefix)])
                                         (detail-line "node is a list symbol")
                                         (hash-set! xml_hash count_sym (add1 (hash-ref xml_hash count_sym 0)))
                                         (let ([new_prefix (string->symbol (format "~a~a" ancester_prefix (hash-ref xml_hash count_sym)))])
                                           (detail-line (format "it is a list sym, prefix changed to:[~a]" new_prefix))
                                           (loop-node new_prefix (cdr xml_list))))
                                       (begin
                                         (detail-line "node is not a list symbol")
                                         (loop-node node (cdr xml_list))))))
                             (begin
                               (let* ([prefix #f]
                                      [node_prefix (car node)]
                                      [attr_list (cadr node)]
                                      [content_list (cddr node)])

                                 (set! prefix (string->symbol (format "~a~a" (if ancester_prefix (format "~a." ancester_prefix) "") node_prefix)))

                                 (detail-line "node is a list")
                                 (detail-line (format "node_prefix:[~a]" node_prefix))
                                 (detail-line (format "prefix:[~a]" prefix))
                                 (detail-line (format "is_list?:[~a]" (hash-has-key? sym_hash prefix)))
                                 (detail-line (format "attrs:[~a]" attr_list))
                                 (detail-line (format "content:[~a]" content_list))

                                 (when (hash-has-key? sym_hash prefix)
                                       (let ([count_sym (format "~a.count" prefix)])
                                         (hash-set! xml_hash count_sym (add1 (hash-ref xml_hash count_sym 0)))
                                         (set! prefix (string->symbol
                                                       (format "~a~a~a"
                                                               (if ancester_prefix (format "~a." ancester_prefix) "")
                                                               node_prefix
                                                               (hash-ref xml_hash count_sym)
                                                               )))
                                         (detail-line (format "it is a list sym, prefix changed to:[~a]" prefix))
                                         ))

                                 (detail-line "process the attrs")
                                 (let loop-attr ([attrs attr_list])
                                   (when (not (null? attrs))
                                         (hash-set! xml_hash (format "~a.~a" prefix (caar attrs)) (cadar attrs))

                                         (loop-attr (cdr attrs))))
                                 (detail-line (format "xml_hash after process attrs:[~a]" xml_hash))

                                 (detail-line "process the content")
                                 (if (null? content_list)
                                     (loop-node prefix '(""))
                                     (loop-node prefix content_list)))

                               (loop-node ancester_prefix (cdr xml_list)))))
                       xml_hash))

                 (detail-new-page)
                 (detail-line "")
                 (detail-line "***************************************")
                 (detail-line (format "xml_hash:[~a]" xml_hash))
                 (detail-line "***************************************")
                 xml_hash))))))))))
