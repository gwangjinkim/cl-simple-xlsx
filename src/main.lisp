(ql:quickload :cl-ppcre)
(ql:quickload :cl-fast-xml)
(ql:quickload :cl-fad)
(ql:quickload :local-time)
(ql:quickload :uiop)
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
	   #:port->lines
	   #:format-w3cdtf
	   #:zip-xlsx
	   #:unzip-xlsx
	   #:date->oa-date-number
	   #:oa-date-number->date
	   #:path-string-make
	   #:get-time-zone
	   ;; #:debug-unzip
	   ))

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

(defun format-w3cdtf (the-date)
  "Format a date object into W3C DTF string format (e.g., 2014-12-15T13:24:27+08:00)."
  (local-time:format-timestring "~Y-~m-~dT~H:~M:~S~z" the-date nil))

(defun format-w3cdtf (the-date)
  "Format a date object into W3C DTF string format (e.g., )."
  (local-time:format-timestring nil the-date :format
				'((:year 4) #\-
				  (:month 2) #\-
				  (:day 2) #\T
				  (:hour 2) #\:
				  (:min 2) #\:
				  (:sec 2) 
				  :gmt-offset
				  )))
;; (:msec 4)

(defun zip-xlsx (zip-file content-dir &optional (content-file-name "\\[Content_Types].xml"))
  "Create a zip archive from the content directory using David Lichteblau's zip library."
  (zip:with-output-to-zipfile (zipwriter zip-file :if-exists :supersede)
    (zip:write-zipentry zipwriter content-file-name
                        (merge-pathnames content-file-name content-dir))
    (dolist (dir '("_rels/" "docProps/" "xl/"))
      (cl-fad:walk-directory (merge-pathnames dir content-dir)
                             (lambda (file)
                               (zip:write-zipentry zipwriter
                                                   (subseq (namestring (merge-pathnames file content-dir)) (length content-dir))
                                                   file))))))


;;;; my correction functions only valid for SBCL

(defun path-to-string (path)
  "Convert Path to plain string expanding `~` correctly."
  (let* ((path-string (format nil "~a" path))
	 (pos (position #\? path-string)))
    (if (and pos (zerop pos))
	(concatenate 'string (uiop:getenv "HOME") (subseq path-string 1))
        path-string)))

(Defun path-string-directory (path &key (sep #\/))
  "Return parent directory string ending with `/`."
  (let* ((path-string (path-to-string path))
	 (pos (position sep (reverse path-string))))
    (if pos
	(subseq path-string 0 (- (length path-string) (position sep (reverse path-string))))
	(uiop/os:getcwd)))) ;; current working directory

(defun path-string-last (path &key (sep #\/))
  "Return last element of a path-string - directory or file."
  (let* ((path-string (path-to-string path))
	 (pos (position sep (reverse path-string))))
    (if pos
	(subseq path-string (- (length path-string) (position sep (reverse path-string))))
	path-string)))

(defun path-string-name (path &key (sep #\/))
  "Return name component of filename (before last dot)."
  (let ((last-part (path-string-last path :sep sep))
        (sep-type #\.))
    (subseq last-part 0 (- (length last-part) (position sep-type (reverse last-part)) 1))))

(defun path-string-type (path &key (sep #\/))
  "Return type portion of filename (after last dot)."
  (let ((last-part (path-string-last path :sep sep))
        (sep-type #\.))
    (subseq last-part (- (length last-part) (position sep-type (reverse last-part))))))

(defun path-string-make (path &key (directory nil) (sep #\/))
  "Return corrected path object of a path or path string by treating brackets as plain
     text."
  (make-pathname :directory (or directory (path-string-directory path :sep sep))
                 :name (path-string-name path :sep sep)
                 :type (path-string-type path :sep sep)))

#|
(defun create-test-file (path &key (text "<H1/>"))
  "If this works without error, square brackets are treated correctly by the system."
  (with-open-file (stream (ensure-directories-exist path)
                          :direction :output
                          :if-does-not-exist :create
                          :if-exists :supersede)
    (format stream "~a" text)))
|#


;;;; This is part of zip package which I need to call to correct the unzip function

(defun %zipfile-entry-contents (entry stream)
  (zip::with-latin1 ()
    (let ((s (zip::zipfile-entry-stream entry))
	  header)
      (file-position s (zip::zipfile-entry-offset entry))
      (setf header (zip::make-local-header s))
      (assert (= (zip::file/signature header) #x04034b50))
      (file-position s (+ (zip::file-position s)
			  (zip::file/name-length header)
			  (zip::file/extra-length header)))
      (let ((in (make-instance 'zip::truncating-stream
			       :input-handle s
			       :size (zip::zipfile-entry-compressed-size entry)))
	    (outbuf nil)
	    out)
	(if stream
	    (setf out stream)
	    (setf outbuf (zip::make-byte-array (zip::zipfile-entry-size entry))
		  out (zip::make-buffer-output-stream outbuf)))
	(ecase (zip::file/method header)
	  (0 (zip::store in out))
	  (8 (zip::inflate in out)))
	outbuf))))

(defun %%zipfile-entry-contents (entry &optional stream)
  (if (pathnamep stream)
      (with-open-file (s (path-string-make stream)
			 :direction :output
			 :if-exists :supersede
                         :element-type '(unsigned-byte 8))
	(%zipfile-entry-contents entry s))
      (%zipfile-entry-contents entry stream)))


(defun better-unzip (pathname target-directory &key (if-exists :error) verbose)
  ;; <Xof> "When reading[1] the value of any pathname component, conforming
  ;;       programs should be prepared for the value to be :unspecific."
  (when (set-difference (list (pathname-name target-directory)
                              (pathname-type target-directory))
                        '(nil :unspecific))
    (error "pathname not a directory, lacks trailing slash?"))
  (zip:with-zipfile (zip pathname)
    (zip:do-zipfile-entries (name entry zip)
      (let ((filename (path-string-make name :directory target-directory)))
        (ensure-directories-exist filename)
        (unless (char= (elt name (1- (length name))) #\/)
          (ecase verbose
            ((nil))
            ((t) (write-string name) (terpri))
            (:dots (write-char #\.)))
          (force-output)
          (with-open-file
              (s filename :direction :output :if-exists if-exists
               :element-type '(unsigned-byte 8))
            (%%zipfile-entry-contents entry s)))))))

#|
(defun debug-unzip (pathname target-directory &key (if-exists :error) verbose)
  (when (set-difference (list (pathname-name target-directory)
			      (pathname-type target-directory))
			'(nil :unspecific))
    (error "pathname not a directory, lacks trailing slash?"))
  (let ((filenames)
	(entries))
    (zip:with-zipfile (zip pathname)
      (zip:do-zipfile-entries (name entry zip)
	(let ((filename name))
	  (push filename filenames)
	  (push entry entries)))
      (list filenames entries))))

("docProps/app.xml" "docProps/core.xml" "xl/sharedStrings.xml" "xl/styles.xml"
 "xl/theme/theme1.xml" "xl/worksheets/sheet2.xml" "xl/worksheets/sheet1.xml"
 "xl/_rels/workbook.xml.rels" "xl/workbook.xml" "_rels/.rels"
"[Content_Types].xml")

This is what name is bearing in this loop!
|#

;;;; Now, unzip can be called.

(defun unzip-xlsx (zip-file content-dir)
  "Unzip a zip archive into the content directory using David Lichteblau's zip library."
  (better-unzip zip-file content-dir :if-exists :supersede))

#|

This corrected unzip corrects with the path-string functions
especially path-string-make the mis-interpretation of [ ] squared brackets
through the Common Lisp pathname system.

By using (format nil "~a" path) it treats pathnames as plain strings.
The danger is of course that the separator (sep #\/) will change in other systems.
thus these functions can pass a different separator anytime.
the separator for name and type part of a filename stays #\. .

By entering extra :directory key argument, one can change the parent directory
when generating via path-string-make.

In case the entire path is just a filename, path-string-directory will return
the current working directory (assuming relative path).
path-string-name and path-string-type will still work on path-string-last, since this will just return the entire path's string.

By applying (path-string-make on the `name` variable of the `zip:do-zipfile-entries` macro, we can avoid that the square brackets get interepreted.

|#



(defun date->oa-date-number (t-date &key (local-time-p t))
  "Convert a date to an OA date number."
  (let* ((epoch (local-time:encode-timestamp 0 0 0 0 30 12 1899
					     :timezone (if local-time-p
							   local-time:*default-timezone*
							   local-time:+utc-zone+)))
         (date-seconds (local-time:timestamp-to-unix t-date)))
    (floor (+ (/ (- date-seconds (local-time:timestamp-to-unix epoch)) 86400.0) 1.0))))

;; (defun oa-date-number->date (oa-date-number &key (local-time-p t))
;;   "Convert an OA date number to a Common Lisp date."
;;   (let* ((epoch (local-time:encode-timestamp 0 0 0 0 30 12 1899
;; 					     :timezone (if local-time-p
;; 							   local-time:*default-timezone*
;; 							   local-time:+gmt-zone+)))
;;          (date-seconds (+ (local-time:timestamp-to-unix epoch)
;; 				 (* (floor oa-date-number)
;; 				    local-time:+seconds-per-day+))))
;;     (local-time:unix-to-timestamp date-seconds)))

(defun oa-date-number->date (oa-date-number &key (local-time-p t))
  "Convert an Excel OA date number to a Common Lisp timestamp using the local-time package."
  (let* ((epoch (local-time:encode-timestamp 0 0 0 0 30 12 1899
                                             :timezone (if local-time-p
                                                           local-time:*default-timezone*
                                                           local-time:+gmt-zone+)))
         (days (floor oa-date-number))
         (fractional-day (- oa-date-number days))
         (seconds (local-time:timestamp-to-unix epoch))
         (date-seconds (+ seconds (* days local-time:+seconds-per-day+)
                          (round (* fractional-day local-time:+seconds-per-day+))))
         (timestamp (local-time:unix-to-timestamp date-seconds)))
    timestamp))

;; no need to calculate to seconds, because we can use
;; local-time:timestamp+ or local-time:timestamp- directly!

;; timezone
;; (local-time:all-timezones-matching-subzone "CEST")
;; the author very likely used "CST" china standard time which is +08:00

(defun get-timezone (&optional (timezone-name "GMT"))
  "Given a timezone-name, return the first of all matching zones."
  (if (string= timezone-name "GMT")
      local-time::+gmt-zone+
      (first (local-time:all-timezones-matching-subzone timezone-name))))


#|
the original Racket code:

(define (date->oa_date_number t_date [local_time? #t])
  (let ([epoch (* -1 (find-seconds 0 0 0 30 12 1899 local_time?))]
        [date_seconds (date->seconds t_date local_time?)])
    (inexact->exact (floor (* (/ (+ date_seconds epoch) 86400000) 1000)))))

(define (oa_date_number->date oa_date_number [local_time? #t])
  (let* ([epoch (* -1 (find-seconds 0 0 0 30 12 1899 local_time?))]
         [date_seconds
          (inexact->exact (floor (- (* (/ (floor oa_date_number) 1000) 86400000) epoch)))]
         [actual_date (seconds->date (+ date_seconds (* 24 60 60)) local_time?)])
(seconds->date (find-seconds 0 0 0 (date-day actual_date) (date-month actual_date) (date-year actual_date) local_time?))))

|#

(defun oa-date-number->date (oa-date-number &key (local-time-p t))
  (let* ((epoch (local-time:timestamp-to-unix
		 (local-time:encode-timestamp 0 0 0 0 30 12 1899
					      :timezone (if local-time-p
							    local-time:*default-timezone*
							    local-time:+gmt-zone+))))
	 (date-seconds (rationalize (floor (+ (* (/ (floor oa-date-number) 1000) 86400000)
					      epoch))))
	 (actual-date (local-time:unix-to-timestamp
		       (+ date-seconds local-time:+seconds-per-day+))))
    actual-date))
								    
;; or one could offer to enter :timezone manually where `get-timezone` would be a help

