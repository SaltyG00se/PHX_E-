(defun C:XFind ( / *error* vars vals ss str objList data)
  (gc)
  
  (prompt "\n")

  ;; Intialize
  (if (= (strcase (getvar "PROGRAM")) "ACAD")
    (princ "Welcome to XFind!")
  )
  ;; Remove this if there are issues with the program
	(if (= (strcase (getvar "PROGRAM")) "ACADLT")
    (princ "False")
  )

  
 ;;Enable modern debugging (AutoCAD 2021+)
  ;; If you edit this code in VS Code, AutoCAD line needs according to: https://help.autodesk.com/view/OARX/2025/ENU/?guid=GUID-037BF4D4-755E-4A5C-8136-80E85CCEDF3E
  (setvar "LISPSYS" 1)  

  ;; Error Handling stuff: check for AutoCAD LT, restore system variables, and exit program under error conditions
  (defun *error* (error)
    (mapcar 'setvar vars vals)  ;; Restore system variables
    (if (not (= (strcase (getvar "PROGRAM")) "ACADLT"))
      (vla-endundomark *doc*)  ;; End undo mark if not AutoCAD LT
    )
    ; Error Handling for exiting program
    (cond
      (not error)
      ((wcmatch (strcase error) "*CANCEL*,*QUIT*")
       (vl-exit-with-error "\r                                              "))
      (1 (vl-exit-with-error (strcat "\r*ERROR*: " error)))
    )
    (princ)
  )

  ;; Initialize AutoCAD Environment 
  (if (/= (strcase (getvar "PROGRAM")) "ACADLT")
    (progn
      ;;Use "vlax-create-object" to create an Excel application object. 
			;;Use "vlax-get-property" to access the "ActiveSheet" property of the Excel application.
			;;Use "vlax-invoke-method" to call a method like "Range" on the active sheet to write data. 
      (setq *acad* (vlax-get-acad-object))
      (setq *doc* (vlax-get *acad* 'ActiveDocument))
      ;; Manage undo marks
      (vla-endundomark *doc*)
      (vla-startundomark *doc*)
    )
  )

  ;; Modify System Variables listed in "vars"
  (setq vars '("cmdecho" "highlight")) ;; (" cmdecho prevents commands from being displayed in the command line" "highlight enables highlighting of selected objects")
  (setq vals (mapcar 'getvar vars))  ;; Save original values
  (mapcar 'setvar vars '(0 1))  ;; Sets cmdecho = 0, highlight = 1

  (and
		;; Get user input for search 
    ;;(possibly na if statement needs to be here. format of if conditions in autolisp below: 
    
		;; (if condition
  	;;(then-expression) ;; Executes if the condition is true
  	;;(else-expression) ;; Executes if the condition is false
    ;;)

		(setq str (getstring T "\nEnter text to find: ")) ;; prompt for user input
    
  	(progn ;; Begins the search program
     
			(setq matchType (strcase (getstring "\nEnter 'E' for exact match or 'P' for partial match: "))) ;; Prompt for match type (exact or partial)
			(if (wcmatch matchType "P") ;;if statment to check if the match type is partial
				(setq str (strcat "*" str "*")) ;; If-P Add wildcards for partial match
				(setq str str) ;; Else-Exact match, no change
			)
			
      ;; Prompt the user to choose layout or entire project
      (setq layoutChoice (strcase (getstring "\nEnter 'L' to search current layout, 'A' for entire project, or 'T' for all tabs: ")))
      ;; Perform the search based on user's choice
      (cond
        ;; Search current layout
        ((wcmatch layoutChoice "L")
          (setq ss (ssget "X" (list '(0 . "*TEXT,MTEXT,ATTRIB") (cons 1 str) (cons 410 (getvar "ctab")))))
        )
        ;; Search entire project (all layouts and model space)
        ((wcmatch layoutChoice "A")
          (setq ss (ssget "X" (list '(0 . "*TEXT,MTEXT,ATTRIB") (cons 1 str))))
        )
        ;; Search all tabs (layouts)
        ((wcmatch layoutChoice "T")
          (progn
            (setq ss nil);; This may be an issue. If so than reduce to, two if statments after debug.
            (vlax-for layout (vla-get-Layouts *doc*)
              (setq ss (append ss (ssget "X" (list '(0 . "*TEXT,MTEXT,ATTRIB") (cons 1 str) (cons 410 (vla-get-Name layout)))))))
        	)
				)
					;; Invalid input, notify user and exit
				(t
					(prompt "\nInvalid choice. Please enter 'L', 'A', or 'T'.")
					
				)
			)
     
    ;; Perform the search for the types of fields listed in the code above (Text, MText, and Attributes). Looking into nested statments in the future.
 		; Check if any objects were found
			(if ss
				(progn
					;; Store found objects in a list
					(setq objList nil row 2)  ;; Start writing from row 2 in Excel
					(repeat (sslength ss)
						(setq obj (ssname ss (setq i (1+ (or i -1)))))
						(setq objList (cons obj objList))
					)

				;; Print object details, possibly can be changed. Source https://www.lee-mac.com/getfieldobjects.html, and CoPilot
					(princ "\nObjects Found:\n-------------------------------------")
					(setq data "")
					(foreach obj objList
						(setq entData (entget obj))
						(setq textVal (cdr (assoc 1 entData)))
						(setq layer (cdr (assoc 8 entData)))
						(setq handle (cdr (assoc 5 entData)))
						(setq objType (cdr (assoc 0 entData)))
						(setq insertion (cdr (assoc 10 entData)))

					

            ;; Formating the Data
            (setq entry (strcat "\nHandle: " handle " | Type: " objType " | Layer: " layer " | Value: " textVal))
            (setq data (strcat data entry))

            ;; Highlight Object
            (redraw obj 3)
          )
					;; Flags:

							;;0: Turns off the highlighting of the object.

							;;1: Turns on the highlighting of the object.

							;;2: Redraws the object without altering its highlighting status.

							;;3: Forces a complete regeneration of the object.

          ;; Display results
          (princ data)

          ;; Keep found items highlighted if this does not turn off adjust flags above or type regen
					;Source: https://www.lee-mac.com/getfieldobjects.html for finding feild objects.
          (sssetfirst nil ss)

          ;; Allow user to trace wires by highlighting found items
          (princ "\nFound objects are highlighted.")
        )
        (princ "\nNo matching text found.")
      )
    )
  )

  (*error* nil)  ;; Restore variables on exit ...
  
)

(defun c:SFind () (c:XFind))  ;; Alias command for XFind