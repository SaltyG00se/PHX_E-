(defun C:XFind ( / *error* vars vals ss str)

   (gc)
   (prompt "\n")
   (defun *error* (error)
      (mapcar 'setvar vars vals)
      (vla-endundomark *doc*)
      (cond
        (not error)
        ((wcmatch (strcase error) "*CANCEL*,*QUIT*")
          (vl-exit-with-error "\r                                              ")
        )
        (1 (vl-exit-with-error (strcat "\r*ERROR*: " error)))
      )
      (princ)
   )
   ;;------------------------------------------
   ;; Intitialze drawing and program variables:
   ;;
   (setq *acad* (vlax-get-acad-object))
   (setq *doc* (vlax-get *acad* 'ActiveDocument))
   (vla-endundomark *doc*)
   (vla-startundomark *doc*)
   (setq vars '("cmdecho" "highlight"))
   (setq vals (mapcar 'getvar vars))
   (mapcar 'setvar vars '(0 1))
   (command "_.expert" (getvar "expert")) ;; dummy command

   (and
     (setq str (getstring T "\nEnter string to find: "))
     (setq str (strcat "*" str "*"))
     (setq ss (ssget "X" (list '(0 . "*TEXT")(cons 1 str)(cons 410 (getvar "ctab")))))
     (princ (strcat "\nFound " (itoa (sslength ss)) " matching text entities."))
     (sssetfirst nil ss)
   )
   (*error* nil)
)
(defun c:SFind ()(c:XFind))