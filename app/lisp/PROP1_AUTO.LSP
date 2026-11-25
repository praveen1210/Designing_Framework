;; =====================================================================
;; PROP1_AUTO.LSP  -  PROP1 table with no manual prompts
;;
;; Same table and same maths as PROP1(SEQ-BODY EXTN).LSP, but the values
;; come from the data file made by the Design Framework web page
;; ("AutoCAD data (.lsp)" button), so 2Tan / L1 / L3 are never typed by hand.
;;
;; Use:
;;   1. APPLOAD this file once (or add it to the Startup Suite).
;;   2. Type PROP1AUTO. If no data is loaded yet it asks for the
;;      PROP1_DATA_*.lsp file, then draws the table at the X,Y in that file.
;;   PROP1PICK = same, but you pick the insertion point on screen.
;;
;; Sheets: the data file says SST-BODY or SST-LEG.
;;   SST-BODY: L2 = BL2 + TAN2*EXT, L1 = L2 - TAN2*HT1, L3 = L2 + TAN2*HT3
;;   SST-LEG : D  = BL2 + TAN2*EXT, L1 = (D - TAN2*HT1)/2, L2 = D/2, L3 = (D + TAN2*HT3)/2
;;   (same as the two Excel sheets)
;;
;; Changes vs PROP1: OSNAP is switched off while drawing (and restored),
;; so snaps can't pull the lines/text to the wrong point.
;; =====================================================================

(defun prop1:get (key) (cdr (assoc key *PROP1-DATA*)))

(defun prop1:ensure-data (/ f)
  (if (not *PROP1-DATA*)
    (if (setq f (getfiled "Select PROP1 data file from Design Framework" "" "lsp" 0))
      (load f)
    )
  )
  (if (and *PROP1-DATA* (prop1:get "BL2") (prop1:get "TAN2") (prop1:get "HT1") (prop1:get "HT3"))
    T
    (progn (princ "\nNo valid PROP1 data loaded.") nil)
  )
)

(defun prop1:text (pt str)
  (command "_TEXT" "_J" "_MC" pt 50 0 str)
)

(defun prop1:draw (IP BL2 TAN2 HT1 HT3 LEG
                   / DD COLW ROWH TOTCOL TOTROW X0 Y0 I ROW YC EXT PAIR
                     L1 L2 L3 D1 D2 D3 old-os old-cmd old-col old-style)
  (setq COLW 2500.0 ROWH 500.0 TOTCOL 7 TOTROW 10)
  (setq old-os (getvar "OSMODE") old-cmd (getvar "CMDECHO")
        old-col (getvar "CECOLOR") old-style (getvar "TEXTSTYLE"))
  (setvar "OSMODE" 0)
  (setvar "CMDECHO" 0)

  (if (not (tblsearch "STYLE" "SIMPLEX"))
    (command "-STYLE" "SIMPLEX" "simplex.shx" "0" "0.8" "0" "N" "N")  ; same answers as PROP1
  )
  (setvar "TEXTSTYLE" "SIMPLEX")
  (setvar "CECOLOR" "6")

  (setq X0 (car IP) Y0 (cadr IP))

  ;; outer border
  (command "_RECTANG" (list X0 Y0) (list (+ X0 (* COLW TOTCOL)) (- Y0 (* ROWH TOTROW))))

  ;; vertical lines
  (setq I 1)
  (repeat (1- TOTCOL)
    (command "_LINE" (list (+ X0 (* I COLW)) Y0) (list (+ X0 (* I COLW)) (- Y0 (* ROWH TOTROW))) "")
    (setq I (1+ I))
  )
  ;; horizontal lines
  (setq I 1)
  (repeat (1- TOTROW)
    (command "_LINE" (list X0 (- Y0 (* I ROWH))) (list (+ X0 (* TOTCOL COLW)) (- Y0 (* I ROWH))) "")
    (setq I (1+ I))
  )

  ;; headers
  (setq I 0)
  (foreach H '("L1" "L1 DIAGONAL" "L2" "L2 DIAGONAL" "L3" "L3 DIAGONAL")
    (prop1:text (list (+ X0 (* COLW (+ 1.5 I))) (- Y0 (/ ROWH 2.0))) H)
    (setq I (1+ I))
  )

  ;; data rows
  (setq ROW 1)
  ;; labels written out (rtos output depends on DIMZIN, so not used here)
  (foreach PAIR '((-3000 . "-3.0M BODY EXTN") (-1500 . "-1.5M BODY EXTN") (0 . "+0.0M BODY EXTN")
                  (1500 . "+1.5M BODY EXTN") (3000 . "+3.0M BODY EXTN") (4500 . "+4.5M BODY EXTN")
                  (6000 . "+6.0M BODY EXTN") (7500 . "+7.5M BODY EXTN") (9000 . "+9.0M BODY EXTN"))
    (setq EXT (car PAIR))
    (setq YC (- Y0 (* ROW ROWH) (/ ROWH 2.0)))
    (prop1:text (list (+ X0 (/ COLW 2.0)) YC) (cdr PAIR))

    (setq DD (+ BL2 (* TAN2 EXT)))
    (if LEG
      (setq L1 (/ (- DD (* TAN2 HT1)) 2.0)
            L2 (/ DD 2.0)
            L3 (/ (+ DD (* TAN2 HT3)) 2.0))
      (setq L1 (- DD (* TAN2 HT1))
            L2 DD
            L3 (+ DD (* TAN2 HT3)))
    )
    (setq D1 (sqrt (* 2.0 L1 L1))
          D2 (sqrt (* 2.0 L2 L2))
          D3 (sqrt (* 2.0 L3 L3)))

    (setq I 0)
    (foreach V (list L1 D1 L2 D2 L3 D3)
      (prop1:text (list (+ X0 (* COLW (+ 1.5 I))) YC) (rtos V 2 2))
      (setq I (1+ I))
    )
    (setq ROW (1+ ROW))
  )

  (setvar "CECOLOR" old-col)
  (setvar "TEXTSTYLE" old-style)
  (setvar "OSMODE" old-os)
  (setvar "CMDECHO" old-cmd)
  (princ (strcat "\nPROP1 table created (" (if LEG "SST-LEG" "SST-BODY") ")."))
)

;; Draw at the X,Y stored in the data file
(defun c:PROP1AUTO (/ x y)
  (if (prop1:ensure-data)
    (progn
      (setq x (prop1:get "X") y (prop1:get "Y"))
      (prop1:draw (list (if x x 0.0) (if y y 0.0) 0.0)
                  (prop1:get "BL2") (prop1:get "TAN2") (prop1:get "HT1") (prop1:get "HT3")
                  (= (prop1:get "SHEET") "SST-LEG"))
    )
  )
  (princ)
)

;; Same data, pick the insertion point on screen
(defun c:PROP1PICK (/ ip)
  (if (and (prop1:ensure-data) (setq ip (getpoint "\nPick table insertion point: ")))
    (prop1:draw ip (prop1:get "BL2") (prop1:get "TAN2") (prop1:get "HT1") (prop1:get "HT3")
                (= (prop1:get "SHEET") "SST-LEG"))
  )
  (princ)
)

;; Forget the loaded data (next PROP1AUTO asks for a file again)
(defun c:PROP1CLEAR () (setq *PROP1-DATA* nil) (princ "\nPROP1 data cleared.") (princ))

(princ "\nPROP1_AUTO loaded: PROP1AUTO (use X,Y from data), PROP1PICK (pick point), PROP1CLEAR.")
(princ)
