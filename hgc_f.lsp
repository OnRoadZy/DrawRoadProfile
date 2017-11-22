(defun C:fhgc ()
  (setq sheet (strcat "Sheet" (itoa (getint "请输入工作表编号：")))
	num-start (getint "请输入起始行：")
	num-end (getint "请输入结束行：")
      	file-name "e:\\gc.xlsx"
	base-point-zh (getpoint "请输入桩号表基准点：")
	base-point-jggc (getpoint "请输入竣工高程表基准点：")
	base-point-zdm (getpoint "请输入纵断面图基准点：")
	base-zh (getreal "请输入基准桩号值：")
	base-gc (getreal "请输入基准高程值："))
  ;取得数据列表：
  (setq zdm-list (get-zdm-list
		   (get-gc-list file-name sheet num-start num-end)))
  ;画图：
  (draw-lczh zdm-list base-point-zh base-zh);画里程桩号
  (draw-jggc zdm-list base-point-jggc base-zh);画竣工高程
  (draw-zdm zdm-list base-point-zdm base-zh base-gc);画纵断面
  )

;取得桩号、高程列表：
(defun get-zdm-list (gc-list)
  (setq len (length gc-list)
	zdm-list nil)
  (setq i (1- len))
  (while (>= i 0)
    (setq lczh (atof (substr (vlax-variant-value (nth 0 (nth i gc-list))) 4))
	  jggc (vlax-variant-value (nth 3 (nth i gc-list))))
    (setq i (1- i))
    (setq zdm-list (cons (cons lczh jggc) zdm-list))))

;;读取文件，获得桩号点和高程值（由exel文件提供）列表：
(defun get-gc-list (file-name sheet num-start num-end)
  (vl-load-com)
  (setq workbooks (get-workbooks (create-app)))
  (setq file-ob (open-file workbooks file-name))
  (setq values
    (get-values
      (create-range-ob
	(get-sheet (get-sheets file-ob) sheet)
	(get-range-str num-start num-end))))
  ;(vlax-invoke-method workbooks "Close");关闭工作薄。
  ;(vlax-invoke-method file-ob "Quit");退出excel对象。
  ;(vlax-release-object file-ob);释放excel对象。
  (values-to-list values))

;读取单元格的值：
(defun get-cell-value (app-ob cell)
    (setq range-str (vlax-get-property app-ob "range" cell))
    (vlax-variant-value (vlax-get-property range-str "Value"))) 
       
;转换为list:
(defun values-to-list (values)
  (vlax-safearray->list (vlax-variant-value values)))

;获取范围对象的值:
(defun get-values (range-ob)
  (vlax-get-property range-ob "Value"))

;用指定的字符串创建工作表范围对象:
(defun create-range-ob (sheet-ob range-str)
  (vlax-get-property sheet-ob "Range" range-str))

;获取指定的工作表:
(defun get-sheet (sheets sheet)
  (vlax-get-property sheets "Item" sheet))

;获取范围字串：(文件排列为：中桩、X、Y、H)
(defun get-range-str (num-start num-end)
  (strcat "A" (itoa num-start) ":D" (itoa num-end)))

;获取工作表集合:
(defun get-sheets (file)
  (vlax-get-property file "Sheets"))

;打开指定的excel文件:
(defun open-file (workbooks file-name)
  (vlax-invoke-method workbooks "open" file-name))

;获取工作薄集合对象:
(defun get-workbooks (app)
  (vlax-get-property app "workbooks"))

;创建程序对象:
(defun create-app ()
  (vlax-get-or-create-object "Excel.Application"))
 
;取得竣工高程值：
(defun get-jggc (zdm-list n)
  (cdr (nth n zdm-list)))

;取得桩号值：
(defun get-zh (zdm-list n)
  (car (nth n zdm-list)))

;取得桩号绘制点：（通用）
(defun get-zh-point (base-point base-zh zdm-list n offset-x offset-y)
  (cons
    (+ (car base-point)
       (- (get-zh zdm-list n) base-zh)
       offset-x)
    (cons
      (+ (car (cdr base-point))
	 offset-y)
      (cdr (cdr base-point)))))

;取得里程桩号绘制点：
(defun get-lczh-point (base-point base-zh zdm-list n)
  (get-zh-point base-point base-zh zdm-list n 6 8))

;取得竣工高程绘制点：
(defun get-jggc-point (base-point base-zh zdm-list n)
  (get-zh-point base-point base-zh zdm-list n 6 2))

;取得纵断面绘制点：
(defun get-zdm-point (base-point base-zh base-gc zdm-list n)
  (get-zh-point base-point base-zh zdm-list n 0
    (* (- (get-jggc zdm-list n) base-gc) 5)))
  
;画里程桩号表：
(defun draw-lczh (zdm-list base-point base-zh)
  (prin1 "现在画里程桩号表……")
  (setq len (length zdm-list)
	i 0)
  (while (< i len)
    (command "text"
	     (get-lczh-point base-point base-zh zdm-list i)
	     "5"
	     "90"
	     (get-zh-str (get-zh zdm-list i))
	     "")
    (setq i (1+ i))))

;整理桩号文字：
(defun get-zh-str (zh)
  (setq zh-str (rtos zh 2 1)
	zero-str "")
  (setq str-len (strlen zh-str))
  (repeat (- 5 str-len) (setq zero-str (strcat "0" zero-str)))
  (strcat "+" zero-str zh-str))

;画竣工高程表：
(defun draw-jggc (zdm-list base-point base-zh)
  (prin1 "现在画竣工高程表……")
  (setq len (length zdm-list)
	i 0)
  (while (< i len)
    (command "text"
	     (get-jggc-point base-point base-zh zdm-list i)
	     "5"
	     "90"
	     (rtos (get-jggc zdm-list i) 2 3)
	     "")
    (setq i (1+ i))))

;画纵断面图：
(defun draw-zdm (zdm-list base-point base-zh base-gc)
  (prin1 "现在画纵断面图……")
  (setq zdm-points (get-zdm-points base-point base-zh base-gc zdm-list))
  (draw-zdm-bzd zdm-points)
  (draw-zdm-lines zdm-points))

;画标注线：
(defun draw-zdm-bz (zdm-point zdm-zh-str zdm-jggc-str)
  (command "qleader"
	   zdm-point
	   "@15<45"
	   "@3<0"
	   ""
	   zdm-zh-str
	   zdm-jggc-str
	   ""))

;画标注点：
(defun draw-zdm-bzd (zdm-points)
  (setq len (length zdm-list)
	i 0)
  (while (< i len)
    (setq zdm-point (nth i zdm-points))
    (setq i (1+ i))
    (command "circle"
	     zdm-point
	     2)))

;;画道路线：
(defun draw-zdm-lines (zdm-points)
  (command "_pline")
  (mapcar 'command zdm-points)
  (command ""))

(defun get-zdm-points (base-point base-zh base-gc zdm-list)
  (setq len (length zdm-list)
	i (1- len)
	zdm-points nil)
  (while (>= i 0)
    (setq zdm-point (get-zdm-point base-point base-zh base-gc zdm-list i))
    (setq i (1- i))
    (setq zdm-points (cons zdm-point zdm-points))))