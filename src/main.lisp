(ql:quickload :cl-ppcre)
(ql:quickload :cl-fast-xml)
(ql:quickload :cl-fad)
(ql:quickload :local-time)
(ql:quickload :zip)

(defpackage cl-simple-xlsx
  (:use :cl :cl-ppcre :cl-fad :local-time :zip :cl-fast-xml)
  (:export #:cell-range-p
	   #:cellp
	   #:row-col->cell
	   #:cell->row-col
	   #:col-abc->number
	   #:col-number->abc
	   #:range->row-col-pair
	   #:*cell-pattern*
	   #:*cell-range-p-pattern*
	   #:*cell-range-pattern*
	   #:*range-xml-pattern*
	   #:*col-range-pattern*
	   #:*row-range-pattern*
	   #:*number-pattern*
	   #:*abc-pattern*
	   #:capacity->range
	   #:range->capacity
	   #:range->range-xml
	   #:range-xml->range
	   #:to-col-range
	   #:to-row-range
	   #:cell-range->cell-list
	   #:get-cell-range-four-sides-cells
	   #:maintain-sheet-data-consistency
	   #:check-lines-p
	   #:check-lines-files-p
	   #:port->lines))

(in-package :cl-simple-xlsx)

(defparameter *cell-pattern* "^([A-Z]+)([0-9]+)$")
(defparameter *cell-range-p-pattern* "^([A-Z]+)([0-9]+)((-|:)([A-Z]+)([0-9]+))*$")
(defparameter *cell-range-pattern* "^([A-Z]+)([0-9]+)(-|:)([A-Z]+)([0-9]+)$")
(defparameter *range-xml-pattern* "^\\$([A-Z]+)\\$([0-9]+):\\$([A-Z]+)\\$([0-9]+)$")
(defparameter *col-range-pattern* "^([0-9]+|[A-Z]+)(-([0-9]+|[A-Z]+)){0,1}$")
(defparameter *number-pattern* "^([0-9]+)$")
(defparameter *abc-pattern* "^([A-Z]+)$")
(defparameter *row-range-pattern* "^([0-9]+)(-([0-9]+)){0,1}$")

(defun cell-range-p (range-string)
  (cl-ppcre:scan *cell-range-p-pattern* (string-upcase range-string)))

(defun cellp (cell-string)
  (cl-ppcre:scan *cell-pattern* (string-upcase cell-string)))

(defun row-col->cell (row col)
  (format nil "~a~a" (col-number->abc col) row))

(defun cell->row-col (cell-string)
  (destructuring-bind ((start-row . start-col) . (end-row . end-col))
      (range->row-col-pair cell-string)
    (declare (ignore end-col end-row))
    (cons start-row start-col)))

(defun col-abc->number (abc)
  (let ((sum 0))
    (dolist (char (coerce (string-upcase abc) 'list) sum)
      (setf sum (+ (* sum 26) (- (char-code char) (char-code #\A) -1))))))

(defun col-number->abc (num)
  (let ((abc ""))
    (loop while (> num 0)
          do (let ((remainder (mod (1- num) 26)))
               (setf abc (concatenate 'string (string (code-char (+ remainder (char-code #\A))))
                                      abc))
               (setf num (floor (1- num) 26))))
    abc))
    

(defun match (string &key pattern)
  (multiple-value-bind (str res) (cl-ppcre:scan-to-strings pattern (string-upcase string))
    (declare (ignore str))
    (when res
      (loop for x across res
	    collect x))))

(defun match-number (string)
  (let ((res (match string :pattern "^([0-9]+)$")))
    (when res
      (parse-integer (car res)))))

(defun match-abc (string)
  (let ((res (match string :pattern "^([A-Z]+)$")))
    (when res
      (car res))))

(defun match-number-or-abc (string)
  (let ((number (match-number string)))
    (if number
	number
	(match-abc string))))

(defun number-or-abc-to-index (x)
  (let ((res (match-number-or-abc x)))
    (cond ((numberp res) res)
	  ((stringp res) (col-abc->number res))
	  (t res))))

(defun match-range (range-string)
  "Takes a cell string or range string like `A1` or `A1:B2`|`A1-B2`
   and returns indexes of columns and rows - always 4 values -
   start-col-index start-row-index end-col-index end-row-index."
  (let* ((result (match range-string :pattern *cell-range-p-pattern*))
	(result (remove-if #'null result)))
    (if (= (length result) 2)
	(destructuring-bind (start-col start-row) result
	  (list (col-abc->number start-col)
		(parse-integer start-row)
		(col-abc->number start-col)
		(parse-integer start-row)))
	(destructuring-bind (start-col start-row second-part delimiter end-col end-row) result
	  (declare (ignore second-part delimiter))
	  (list (col-abc->number start-col)
		(parse-integer start-row)
		(col-abc->number end-col)
		(parse-integer end-row))))))

(defun range->row-col-pair (range-string)
  (let ((result (match-range range-string)))
    (destructuring-bind (start-col-index
			 start-row-index
			 end-col-index
			 end-row-index) result
      (if (and (>= end-row-index start-row-index)
	       (>= end-col-index start-col-index))
	  (cons (cons start-row-index start-col-index)
		(cons end-row-index end-col-index))
	  '((1 . 1) . (1 . 1))))))

(defun range->capacity (range)
  (destructuring-bind ((start-row . start-col) . (end-row . end-col)) (range->row-col-pair range)
    (cons (1+ (- end-row start-row))
	  (1+ (- end-col start-col)))))

(defun capacity->range (capacity &optional (start-cell "A1"))
  (destructuring-bind (start-row . start-col) (cell->row-col start-cell)
    (destructuring-bind (delta-row . delta-col) capacity 
      (format nil "~a:~a~a"
	      start-cell
	      (col-number->abc (1- (+ start-col delta-col)))
	      (1- (+ start-row delta-row))))))

(defun range->range-xml (range-string)
  (destructuring-bind (start-col-name
		       start-row-index
		       second-part
		       delimiter
		       end-col-name
		       end-row-index) (match range-string :pattern *cell-range-p-pattern*)
    (declare (ignore second-part delimiter))
    (format nil "$~a$~a:$~a$~a"
	    start-col-name
	    start-row-index
	    end-col-name
	    end-row-index)))

(defun range-xml->range (range-xml-string)
  (destructuring-bind (start-col-name
		       start-row-index
		       end-col-name
		       end-row-index) (match range-xml-string :pattern *range-xml-pattern*)
    (format nil "~a~a-~a~a"
	    start-col-name
	    start-row-index
	    end-col-name
	    end-row-index)))

(defun %to-range (range-string &key pattern)
  (let* ((res (match range-string :pattern pattern))
	 (res (remove-if #'null res)))
    (if res
	(cond ((= (length res) 3)
	       (destructuring-bind (start-col second-part end-col) res
		 (declare (ignore second-part))
		 (let ((start-col-index (number-or-abc-to-index start-col))
		       (end-col-index (number-or-abc-to-index end-col)))
		   (if (<= start-col-index end-col-index)
		       (cons start-col-index end-col-index)
		       (cons 1 1)))))
	      ((= (length res) 1)
	       (let ((index (number-or-abc-to-index (car res))))
		 (cons index index)))
	      (t (cons 1 1)))
	(cons 1 1))))


(defun to-col-range (col-range)
  (%to-range col-range :pattern *col-range-pattern*))

(defun to-row-range (row-range)
  (%to-range row-range :pattern *row-range-pattern*))

	       
;; (remove-if #'null (match "A-B-C" :pattern *col-range-pattern*))

(defun cell-range->cell-list (range-string)
  (destructuring-bind ((start-row-index . start-col-index) . (end-row-index . end-col-index))
      (range->row-col-pair range-string)
    (let ((cell-list '()))
      (loop for row-index from start-row-index to end-row-index
	    do (loop for col-index from start-col-index to end-col-index
		     do (push (format nil "~a~a"
				      (col-number->abc col-index)
				      row-index)
			      cell-list)))
      (nreverse cell-list))))

(defun get-cell-range-four-sides-cells (cell-range-string)
  (destructuring-bind ((start-row-index . start-col-index) . (end-row-index . end-col-index))
      (range->row-col-pair cell-range-string)
    (let ((left-cells '())
	  (right-cells '())
	  (top-cells '())
	  (bottom-cells '()))
      (loop for cell in (cell-range->cell-list cell-range-string)
	    do (destructuring-bind (row . col) (cell->row-col cell)
		 (when (= start-col-index col)
		   (push cell left-cells))
		 (when (= end-col-index col)
		   (push cell right-cells))
		 (when (= start-row-index row)
		   (push cell top-cells))
		 (when (= end-row-index row)
		   (push cell bottom-cells)))
	    finally (return (values
			     (nreverse left-cells)
			     (nreverse right-cells)
			     (nreverse top-cells)
			     (nreverse bottom-cells)))))))
	
;; lib.rkt

(defun maintain-sheet-data-consistency (data-list pad-fill)
  (when (null data-list)
    (error "data list is empty"))

  (let ((max-child-length (reduce (lambda (a b) (max a b))
				  (mapcar (lambda (row) (length row))
					  data-list))))
    (mapcar
     (lambda (row)
       (if (< (length row) max-child-length)
	   (append row (make-list (- max-child-length (length row)) :initial-element pad-fill))
	   row))
     data-list)))




(defun port->lines (stream)
  (loop for line = (read-line stream nil nil)
        while line
        collect line))

(defun check-lines-p (expected-port test-port)
  (let* ((expected-lines (port->lines expected-port))
	 (test-lines (port->lines test-port))
	 (test-length (length test-lines)))
    (if (= (length expected-lines) 0)
	(when (not (zerop test-length))
	  (error (format nil "error! expect no content, but actual have [~a] lines" test-length)))
	(loop for line in expected-lines
	      for line-no from 0
	      do (cond ((>= line-no test-length)
			(error (format nil "error! line[~a] expected:[~a] actual:null"
				       (1+ line-no)
				       line)))
		       ((string/= line (elt test-lines line-no))
			(error (format nil "error! line[~a] expected:[~a] actual:[~a]"
				       (1+ line-no)
				       line
				       (elt test-lines line-no))))))))
  t)
				       
(defun check-lines-files-p (expected-file-path test-file-path)
  (with-open-file (expected-port expected-file-path :direction :input)
    (with-open-file (test-port test-file-path :direction :input)
      (check-lines-p expected-port test-port))))

;; (defun format-w3cdtf (the-date)
;;   "Format a date object into W3C DTF string format (e.g., 2014-12-15T13:24:27+08:00)."
;;   (local-time:format-timestring "~Y-~m-~dT~H:~M:~S~z" the-date nil))

(defun timezone-offset (timestamp)
  "Calculate the timezone offset in seconds for the given timestamp."
  (let* ((local-time (local-time:encode-timestamp 
                      (local-time:timestamp-second timestamp)
                      (local-time:timestamp-minute timestamp)
                      (local-time:timestamp-hour timestamp)
                      (local-time:timestamp-day timestamp)
                      (local-time:timestamp-month timestamp)
                      (local-time:timestamp-year timestamp)
                      :timezone (local-time:timestamp-timezone timestamp)))
         (utc-time (local-time:encode-timestamp 
                    (local-time:timestamp-second timestamp)
                    (local-time:timestamp-minute timestamp)
                    (local-time:timestamp-hour timestamp)
                    (local-time:timestamp-day timestamp)
                    (local-time:timestamp-month timestamp)
                    (local-time:timestamp-year timestamp)
                    :timezone (local-time:find-timezone-by-location-name "UTC"))))
    (- (local-time:timestamp-to-unix local-time)
       (local-time:timestamp-to-unix utc-time))))

(defun format-w3cdtf (the-date)
  "Format a date object into W3C DTF string format (e.g., 2014-12-15T13:24:27+08:00)."
  (let* ((year (format nil "~4,'0d" (local-time:timestamp-year the-date)))
         (month (format nil "~2,'0d" (local-time:timestamp-month the-date)))
         (day (format nil "~2,'0d" (local-time:timestamp-day the-date)))
         (hour (format nil "~2,'0d" (local-time:timestamp-hour the-date)))
         (minute (format nil "~2,'0d" (local-time:timestamp-minute the-date)))
         (second (format nil "~2,'0d" (local-time:timestamp-second the-date)))
         (tz-offset (timezone-offset the-date))  ; in seconds
         (tz-sign (if (>= tz-offset 0) "+" "-"))
         (tz-hours (format nil "~2,'0d" (floor (/ (abs tz-offset) 3600)))))
    (format nil "~a-~a-~aT~a:~a:~a~a~a:00"
            year month day hour minute second tz-sign tz-hours)))

(defun zip-xlsx (zip-file content-dir)
  "Create a zip archive from the content directory using David Lichteblau's zip library."
  (zip:with-output-to-zipfile (zipwriter zip-file :if-exists :supersede)
    (zip:write-zipentry zipwriter "[Content_Types].xml"
                        (merge-pathnames "[Content_Types].xml" content-dir))
    (dolist (dir '("_rels/" "docProps/" "xl/"))
      (cl-fad:walk-directory (merge-pathnames dir content-dir)
                             (lambda (file)
                               (zip:write-zipentry zipwriter
                                                   (subseq (namestring (merge-pathnames file content-dir)) (length content-dir))
                                                   file))))))

(defun unzip-xlsx (zip-file content-dir)
  "Unzip a zip archive into the content directory using David Lichteblau's zip library."
  (zip:unzip zip-file content-dir :if-exists :error))

(defun date->oa-date-number (t-date)
  "Convert a date to an OA date number."
  (let* ((epoch (local-time:encode-timestamp 0 0 0 0 30 12 1899 :timezone :utc))
         (date-seconds (local-time:timestamp-to-unix t-date)))
    (floor (+ (/ (- date-seconds (local-time:timestamp-to-unix epoch)) 86400.0) 1.0))))

(defun oa-date-number->date (oa-date-number)
  "Convert an OA date number to a Common Lisp date."
  (let* ((epoch (local-time:encode-timestamp 0 0 0 0 30 12 1899 :timezone :utc))
         (date-seconds (+ (local-time:timestamp-to-unix epoch) (* oa-date-number 86400.0))))
    (local-time:unix-to-timestamp date-seconds)))
