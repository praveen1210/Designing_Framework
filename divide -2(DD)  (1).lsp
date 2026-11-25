(defun c:DD (/ src txt val res)

  (setq src (car (entsel "\nSelect source TEXT: ")))

  (if (and src (= (cdr (assoc 0 (entget src))) "TEXT"))
    (progn
      (setq txt (cdr (assoc 1 (entget src))))
      (setq val (atof txt))
      (setq res (rtos (/ val 2.0) 2 2))

      (prompt (strcat "\nResult = " res))

      ;; Place result in selected text
      (setq dst (car (entsel "\nSelect destination TEXT: ")))
      (if dst
        (entmod
          (subst
            (cons 1 res)
            (assoc 1 (entget dst))
            (entget dst)
          )
        )
      )
    )
  )
  (princ)
)
