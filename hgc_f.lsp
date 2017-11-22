(defun C:fhgc ()
  (setq sheet (strcat "Sheet" (itoa (getint "�����빤�����ţ�")))
	num-start (getint "��������ʼ�У�")
	num-end (getint "����������У�")
      	file-name "e:\\gc.xlsx"
	base-point-zh (getpoint "������׮�ű��׼�㣺")
	base-point-jggc (getpoint "�����뿢���̱߳��׼�㣺")
	base-point-zdm (getpoint "�������ݶ���ͼ��׼�㣺")
	base-zh (getreal "�������׼׮��ֵ��")
	base-gc (getreal "�������׼�߳�ֵ��"))
  ;ȡ�������б�
  (setq zdm-list (get-zdm-list
		   (get-gc-list file-name sheet num-start num-end)))
  ;��ͼ��
  (draw-lczh zdm-list base-point-zh base-zh);�����׮��
  (draw-jggc zdm-list base-point-jggc base-zh);�������߳�
  (draw-zdm zdm-list base-point-zdm base-zh base-gc);���ݶ���
  )

;ȡ��׮�š��߳��б�
(defun get-zdm-list (gc-list)
  (setq len (length gc-list)
	zdm-list nil)
  (setq i (1- len))
  (while (>= i 0)
    (setq lczh (atof (substr (vlax-variant-value (nth 0 (nth i gc-list))) 4))
	  jggc (vlax-variant-value (nth 3 (nth i gc-list))))
    (setq i (1- i))
    (setq zdm-list (cons (cons lczh jggc) zdm-list))))

;;��ȡ�ļ������׮�ŵ�͸߳�ֵ����exel�ļ��ṩ���б�
(defun get-gc-list (file-name sheet num-start num-end)
  (vl-load-com)
  (setq workbooks (get-workbooks (create-app)))
  (setq file-ob (open-file workbooks file-name))
  (setq values
    (get-values
      (create-range-ob
	(get-sheet (get-sheets file-ob) sheet)
	(get-range-str num-start num-end))))
  ;(vlax-invoke-method workbooks "Close");�رչ�������
  ;(vlax-invoke-method file-ob "Quit");�˳�excel����
  ;(vlax-release-object file-ob);�ͷ�excel����
  (values-to-list values))

;��ȡ��Ԫ���ֵ��
(defun get-cell-value (app-ob cell)
    (setq range-str (vlax-get-property app-ob "range" cell))
    (vlax-variant-value (vlax-get-property range-str "Value"))) 
       
;ת��Ϊlist:
(defun values-to-list (values)
  (vlax-safearray->list (vlax-variant-value values)))

;��ȡ��Χ�����ֵ:
(defun get-values (range-ob)
  (vlax-get-property range-ob "Value"))

;��ָ�����ַ�������������Χ����:
(defun create-range-ob (sheet-ob range-str)
  (vlax-get-property sheet-ob "Range" range-str))

;��ȡָ���Ĺ�����:
(defun get-sheet (sheets sheet)
  (vlax-get-property sheets "Item" sheet))

;��ȡ��Χ�ִ���(�ļ�����Ϊ����׮��X��Y��H)
(defun get-range-str (num-start num-end)
  (strcat "A" (itoa num-start) ":D" (itoa num-end)))

;��ȡ��������:
(defun get-sheets (file)
  (vlax-get-property file "Sheets"))

;��ָ����excel�ļ�:
(defun open-file (workbooks file-name)
  (vlax-invoke-method workbooks "open" file-name))

;��ȡ���������϶���:
(defun get-workbooks (app)
  (vlax-get-property app "workbooks"))

;�����������:
(defun create-app ()
  (vlax-get-or-create-object "Excel.Application"))
 
;ȡ�ÿ����߳�ֵ��
(defun get-jggc (zdm-list n)
  (cdr (nth n zdm-list)))

;ȡ��׮��ֵ��
(defun get-zh (zdm-list n)
  (car (nth n zdm-list)))

;ȡ��׮�Ż��Ƶ㣺��ͨ�ã�
(defun get-zh-point (base-point base-zh zdm-list n offset-x offset-y)
  (cons
    (+ (car base-point)
       (- (get-zh zdm-list n) base-zh)
       offset-x)
    (cons
      (+ (car (cdr base-point))
	 offset-y)
      (cdr (cdr base-point)))))

;ȡ�����׮�Ż��Ƶ㣺
(defun get-lczh-point (base-point base-zh zdm-list n)
  (get-zh-point base-point base-zh zdm-list n 6 8))

;ȡ�ÿ����̻߳��Ƶ㣺
(defun get-jggc-point (base-point base-zh zdm-list n)
  (get-zh-point base-point base-zh zdm-list n 6 2))

;ȡ���ݶ�����Ƶ㣺
(defun get-zdm-point (base-point base-zh base-gc zdm-list n)
  (get-zh-point base-point base-zh zdm-list n 0
    (* (- (get-jggc zdm-list n) base-gc) 5)))
  
;�����׮�ű�
(defun draw-lczh (zdm-list base-point base-zh)
  (prin1 "���ڻ����׮�ű���")
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

;����׮�����֣�
(defun get-zh-str (zh)
  (setq zh-str (rtos zh 2 1)
	zero-str "")
  (setq str-len (strlen zh-str))
  (repeat (- 5 str-len) (setq zero-str (strcat "0" zero-str)))
  (strcat "+" zero-str zh-str))

;�������̱߳�
(defun draw-jggc (zdm-list base-point base-zh)
  (prin1 "���ڻ������̱߳���")
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

;���ݶ���ͼ��
(defun draw-zdm (zdm-list base-point base-zh base-gc)
  (prin1 "���ڻ��ݶ���ͼ����")
  (setq zdm-points (get-zdm-points base-point base-zh base-gc zdm-list))
  (draw-zdm-bzd zdm-points)
  (draw-zdm-lines zdm-points))

;����ע�ߣ�
(defun draw-zdm-bz (zdm-point zdm-zh-str zdm-jggc-str)
  (command "qleader"
	   zdm-point
	   "@15<45"
	   "@3<0"
	   ""
	   zdm-zh-str
	   zdm-jggc-str
	   ""))

;����ע�㣺
(defun draw-zdm-bzd (zdm-points)
  (setq len (length zdm-list)
	i 0)
  (while (< i len)
    (setq zdm-point (nth i zdm-points))
    (setq i (1+ i))
    (command "circle"
	     zdm-point
	     2)))

;;����·�ߣ�
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