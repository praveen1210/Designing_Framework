(defun c:PROP1
 (/ ent BL2 TAN2 HT1 HT3 IP
    COLW ROWH TOTCOL TOTROW
    P1 P2 P3 P4
    EXT ROW YC
    L1 L2 L3 D1 D2 D3
    X0 Y0
    I)

 (setq ent (car (entsel "\nSelect L2 Source Text : ")))

 (if ent
  (progn

   (setq BL2 (atof (cdr (assoc 1 (entget ent)))))

   (setq TAN2 (getreal "\nEnter 2Tan Value : "))
   (setq HT1  (getreal "\nEnter L1 Height : "))
   (setq HT3  (getreal "\nEnter L3 Height : "))

   (setq IP (getpoint "\nPick Table Insertion Point : "))

   ;; TABLE SIZE

   (setq COLW 2500.0)
   (setq ROWH 500.0)

   ;; 1 Title Column + 6 Data Columns

   (setq TOTCOL 7)

   ;; 1 Header Row + 9 Data Rows

   (setq TOTROW 10)

   ;; SIMPLEX STYLE

   (if (not (tblsearch "STYLE" "SIMPLEX"))
    (command "-STYLE" "SIMPLEX" "simplex.shx" "0" "0.8" "0" "N" "N")
   )

   (setvar "TEXTSTYLE" "SIMPLEX")
   (setvar "CECOLOR" "6")

   (setq X0 (car IP))
   (setq Y0 (cadr IP))

   ;; ====================================
   ;; OUTER BORDER
   ;; ====================================

   (setq P1 IP)
   (setq P2 (list (+ X0 (* COLW TOTCOL)) Y0))
   (setq P3 (list (+ X0 (* COLW TOTCOL))
                  (- Y0 (* ROWH TOTROW))))
   (setq P4 (list X0
                  (- Y0 (* ROWH TOTROW))))

   (command "LINE" P1 P2 "")
   (command "LINE" P2 P3 "")
   (command "LINE" P3 P4 "")
   (command "LINE" P4 P1 "")

   ;; ====================================
   ;; VERTICAL LINES
   ;; ====================================

   (setq I 1)

   (repeat (1- TOTCOL)

    (command
      "LINE"
      (list (+ X0 (* I COLW)) Y0)
      (list (+ X0 (* I COLW))
            (- Y0 (* ROWH TOTROW)))
      ""
    )

    (setq I (1+ I))
   )

   ;; ====================================
   ;; HORIZONTAL LINES
   ;; ====================================

   (setq I 1)

   (repeat (1- TOTROW)

    (command
      "LINE"
      (list X0 (- Y0 (* I ROWH)))
      (list (+ X0 (* TOTCOL COLW))
            (- Y0 (* I ROWH)))
      ""
    )

    (setq I (1+ I))
   )

   ;; ====================================
   ;; COLUMN HEADERS
   ;; ====================================

   (command "_TEXT" "_J" "_MC"
            (list (+ X0 (* COLW 1.5))
                  (- Y0 (/ ROWH 2)))
            50 0 "L1")

   (command "_TEXT" "_J" "_MC"
            (list (+ X0 (* COLW 2.5))
                  (- Y0 (/ ROWH 2)))
            50 0 "L1 DIAGONAL")

   (command "_TEXT" "_J" "_MC"
            (list (+ X0 (* COLW 3.5))
                  (- Y0 (/ ROWH 2)))
            50 0 "L2")

   (command "_TEXT" "_J" "_MC"
            (list (+ X0 (* COLW 4.5))
                  (- Y0 (/ ROWH 2)))
            50 0 "L2 DIAGONAL")

   (command "_TEXT" "_J" "_MC"
            (list (+ X0 (* COLW 5.5))
                  (- Y0 (/ ROWH 2)))
            50 0 "L3")

   (command "_TEXT" "_J" "_MC"
            (list (+ X0 (* COLW 6.5))
                  (- Y0 (/ ROWH 2)))
            50 0 "L3 DIAGONAL")

   ;; ====================================
   ;; DATA ROWS
   ;; ====================================

   (setq ROW 1)

   (foreach EXT
    '(-3000 -1500 0 1500 3000 4500 6000 7500 9000)

    (setq YC
          (- Y0
             (* ROW ROWH)
             (/ ROWH 2.0)))

    ;; Row Title

    (command "_TEXT" "_J" "_MC"
      (list (+ X0 (/ COLW 2.0)) YC)
      50 0
      (cond
       ((= EXT -3000) "-3.0M BODY EXTN")
       ((= EXT -1500) "-1.5M BODY EXTN")
       ((= EXT 0)     "+0.0M BODY EXTN")
       ((= EXT 1500)  "+1.5M BODY EXTN")
       ((= EXT 3000)  "+3.0M BODY EXTN")
       ((= EXT 4500)  "+4.5M BODY EXTN")
       ((= EXT 6000)  "+6.0M BODY EXTN")
       ((= EXT 7500)  "+7.5M BODY EXTN")
       (T             "+9.0M BODY EXTN")
      )
    )

    ;; Calculations

    (setq L2 (+ BL2 (* TAN2 EXT)))

    (setq L1 (- L2 (* TAN2 HT1)))
    (setq L3 (+ L2 (* TAN2 HT3)))

    (setq D1 (sqrt (+ (* L1 L1) (* L1 L1))))
    (setq D2 (sqrt (+ (* L2 L2) (* L2 L2))))
    (setq D3 (sqrt (+ (* L3 L3) (* L3 L3))))

    ;; L1

    (command "_TEXT" "_J" "_MC"
      (list (+ X0 (* COLW 1.5)) YC)
      50 0 (rtos L1 2 2))

    ;; L1 DIAGONAL

    (command "_TEXT" "_J" "_MC"
      (list (+ X0 (* COLW 2.5)) YC)
      50 0 (rtos D1 2 2))

    ;; L2

    (command "_TEXT" "_J" "_MC"
      (list (+ X0 (* COLW 3.5)) YC)
      50 0 (rtos L2 2 2))

    ;; L2 DIAGONAL

    (command "_TEXT" "_J" "_MC"
      (list (+ X0 (* COLW 4.5)) YC)
      50 0 (rtos D2 2 2))

    ;; L3

    (command "_TEXT" "_J" "_MC"
      (list (+ X0 (* COLW 5.5)) YC)
      50 0 (rtos L3 2 2))

    ;; L3 DIAGONAL

    (command "_TEXT" "_J" "_MC"
      (list (+ X0 (* COLW 6.5)) YC)
      50 0 (rtos D3 2 2))

    (setq ROW (1+ ROW))
   )

   (setvar "CECOLOR" "BYLAYER")

   (princ "\nPROP1 Table Created Successfully.")
  )
 )

 (princ)
)