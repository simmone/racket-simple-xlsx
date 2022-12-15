#lang racket

(require rackunit/text-ui rackunit)

(require"../../style/style.rkt")
(require"../../style/font-style.rkt")
(require"../../style/fill-style.rkt")
(require"../../style/alignment-style.rkt")
(require"../../style/border-style.rkt")
(require"../../style/number-style.rkt")

(define test-style
  (test-suite
   "test-style"

   (test-case
    "test-style"

    (let ([style (style-from-hash-code "<s>10<p>Arial<p>0000ff<s><s><s>")])
      (check-false (STYLE-border_style style))
      (check-false (STYLE-alignment_style style))
      (check-false (STYLE-number_style style))
      (check-false (STYLE-fill_style style))
      (check-equal? (FONT-STYLE-hash_code (STYLE-font_style style)) "10<p>Arial<p>0000FF")
      (check-equal? (STYLE-hash_code style) "<s>10<p>Arial<p>0000FF<s><s><s>"))

    (let ([style (style-from-hash-code "<s><s><s>0.000<s>")])
      (check-false (STYLE-border_style style))
      (check-false (STYLE-alignment_style style))
      (check-false (STYLE-font_style style))
      (check-false (STYLE-fill_style style))
      (check-equal? (NUMBER-STYLE-hash_code (STYLE-number_style style)) "0.000")
      (check-equal? (STYLE-hash_code style) "<s><s><s>0.000<s>"))

    (let ([style (style-from-hash-code "<s><s><s><s>FF0000<p>solid")])
      (check-false (STYLE-border_style style))
      (check-false (STYLE-alignment_style style))
      (check-false (STYLE-font_style style))
      (check-false (STYLE-number_style style))
      (check-equal? (FILL-STYLE-hash_code (STYLE-fill_style style)) "FF0000<p>solid")
      (check-equal? (STYLE-hash_code style) "<s><s><s><s>FF0000<p>solid"))

    (let ([style (style-from-hash-code "<s><s>center<p>center<s><s>")])
      (check-false (STYLE-border_style style))
      (check-false (STYLE-fill_style style))
      (check-false (STYLE-font_style style))
      (check-false (STYLE-number_style style))
      (check-equal? (ALIGNMENT-STYLE-hash_code (STYLE-alignment_style style)) "center<p>center")
      (check-equal? (STYLE-hash_code style) "<s><s>center<p>center<s><s>"))

    (let ([style (style-from-hash-code "f00000<p>thin<p>0f0000<p>thick<p>00f000<p>double<p>000f00<p>dashed<s><s><s><s>")])
      (check-false (STYLE-alignment_style style))
      (check-false (STYLE-fill_style style))
      (check-false (STYLE-font_style style))
      (check-false (STYLE-number_style style))
      (check-equal? (BORDER-STYLE-hash_code (STYLE-border_style style)) "F00000<p>thin<p>0F0000<p>thick<p>00F000<p>double<p>000F00<p>dashed")
      (check-equal? (STYLE-hash_code style) "F00000<p>thin<p>0F0000<p>thick<p>00F000<p>double<p>000F00<p>dashed<s><s><s><s>"))

    (let ([style
              (style-from-hash-code
               "f00000<p>thin<p>0f0000<p>thick<p>00f000<p>double<p>000f00<p>dashed<s>10<p>Arial<p>0000ff<s>center<p>center<s>0.000<s>FF0000<p>solid")])
      (check-equal? (ALIGNMENT-STYLE-hash_code (STYLE-alignment_style style)) "center<p>center")
      (check-equal? (FILL-STYLE-hash_code (STYLE-fill_style style)) "FF0000<p>solid")
      (check-equal? (FONT-STYLE-hash_code (STYLE-font_style style)) "10<p>Arial<p>0000FF")
      (check-equal? (NUMBER-STYLE-hash_code (STYLE-number_style style)) "0.000")
      (check-equal? (BORDER-STYLE-hash_code (STYLE-border_style style)) "F00000<p>thin<p>0F0000<p>thick<p>00F000<p>double<p>000F00<p>dashed")
      (check-equal?
       (STYLE-hash_code style)
       "F00000<p>thin<p>0F0000<p>thick<p>00F000<p>double<p>000F00<p>dashed<s>10<p>Arial<p>0000FF<s>center<p>center<s>0.000<s>FF0000<p>solid"))
    )
   ))

(run-tests test-style)
