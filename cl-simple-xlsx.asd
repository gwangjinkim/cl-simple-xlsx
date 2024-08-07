(defsystem "cl-simple-xlsx"
  :version "0.0.1"
  :author "Gwang-Jin Kim"
  :mailto "gwang.jin.kim.phd@gmail.com"
  :license "MIT"
  :depends-on ("cl-ppcre"
               "cl-fast-xml"
               "alexandria")
  :components ((:module "src"
                :components
                ((:file "main"))))
  :description "Read-write .xlsx files in Common Lisp. Translation of the racket-simple-xlsx package
source https://github.com/simmone/racket-simple-xlsx by Chen Xiao. 
Translated (ongoing) by Gwang-Jin Kim <gwang.jin.kim.phd@gmail.com>"
  :in-order-to ((test-op (test-op "cl-simple-xlsx/tests"))))

(defsystem "cl-simple-xlsx/tests"
  :author "Gwang-Jin Kim"
  :license "MIT"
  :depends-on ("cl-simple-xlsx"
               "rove")
  :components ((:module "tests"
                :components
                ((:file "main"))))
  :description "Test system for cl-simple-xlsx"
  :perform (test-op (op c) (symbol-call :rove :run c)))
