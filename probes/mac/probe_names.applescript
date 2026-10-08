-- probe_names.applescript — P0 probe: how much of the Name Manager can AppleScript drive?
--
--   osascript probe_names.applescript <save-folder-POSIX>
--
-- Works on a NEW blank workbook only; saves a copy into <save-folder> so the wrapper can
-- read what Excel stored in xl/workbook.xml, then closes it unsaved. Every test is in its
-- own try block: a FAIL is a result, not a crash. Output: "Txx | STATUS | detail" lines.
-- Test IDs match probes/windows/probe_names.ps1 and probes/officescripts/probe_names.ts.

global rpt

on addLine(tid, verdict, det)
	set rpt to rpt & tid & " | " & verdict & " | " & det & linefeed
end addLine

on repeatText(s, n)
	set out to ""
	repeat n times
		set out to out & s
	end repeat
	return out
end repeatText

on run argv
	set rpt to ""
	set saveDir to item 1 of argv
	set lam to character id 955
	set nl to linefeed

	tell application "Microsoft Excel"
		activate
		set wb to make new workbook
		set s1 to worksheet 1 of wb
		set name of s1 to "S1"
		try
			set s2 to make new worksheet at end of wb
		on error
			set s2 to make new worksheet at wb
		end try
		set name of s2 to "S2"
		set calculation to calculation automatic

		-- T00 environment -------------------------------------------------------------
		try
			my addLine("T00", "INFO", "Excel version " & (version as text))
		end try

		-- T01 workbook-scoped constant ------------------------------------------------
		try
			make new named item at wb with properties {name:"P_K", references:"=10"}
			my addLine("T01", "PASS", "created P_K; references=" & (references of named item "P_K" of wb))
		on error e
			my addLine("T01", "FAIL", e)
		end try

		-- T02 LAMBDA name, called from a cell -----------------------------------------
		try
			make new named item at wb with properties {name:"P_Add1", references:"=LAMBDA(x, x+1)"}
			set formula2 of range "A1" of s1 to "=P_Add1(41)"
			calculate
			set v to value of range "A1" of s1
			if (v as text) is "42" or (v as text) is "42.0" or (v as text) is "42,0" then
				my addLine("T02", "PASS", "=P_Add1(41) -> " & (v as text))
			else
				my addLine("T02", "FAIL", "=P_Add1(41) -> " & (v as text))
			end if
		on error e
			my addLine("T02", "FAIL", e)
		end try

		-- T03 stored text, English vs local --------------------------------------------
		try
			set ni to named item "P_Add1" of wb
			my addLine("T03", "INFO", "references=" & (references of ni))
			my addLine("T03", "INFO", "reference local=" & (reference local of ni))
		on error e
			my addLine("T03", "FAIL", e)
		end try

		-- T04 Name Manager comment (not in the dictionary: expected FAIL) ---------------
		try
			set comment of named item "P_Add1" of wb to "probe comment"
			my addLine("T04", "PASS?", "set comment did not error; read back: " & (comment of named item "P_Add1" of wb))
		on error e
			my addLine("T04", "FAIL", "comment: " & e)
		end try

		-- T05 sheet-scoped names --------------------------------------------------------
		try
			make new named item at s2 with properties {name:"P_Local", references:"=5"}
			set formula2 of range "A1" of s2 to "=P_Local"
			set formula2 of range "A3" of s1 to "=S2!P_Local"
			calculate
			my addLine("T05a", "INFO", "make at sheet: S2!A1=" & ((value of range "A1" of s2) as text) & " S1!A3=" & ((value of range "A3" of s1) as text))
		on error e
			my addLine("T05a", "FAIL", e)
		end try
		try
			make new named item at wb with properties {name:"S2!P_Local2", references:"=6"}
			my addLine("T05b", "INFO", "make at wb with 'S2!' prefix: ok")
		on error e
			my addLine("T05b", "FAIL", e)
		end try
		try
			set nms to ""
			repeat with i from 1 to (count of named items of s2)
				set nms to nms & (name of named item i of s2) & "; "
			end repeat
			my addLine("T05c", "INFO", "named items of S2: " & nms)
		on error e
			my addLine("T05c", "FAIL", e)
		end try

		-- T06 name over a spilled range (#) ------------------------------------------
		try
			set formula2 of range "B1" of s1 to "=SEQUENCE(1,5)"
			make new named item at wb with properties {name:"P_Spill", references:"=S1!$B$1#"}
			set formula2 of range "C2" of s1 to "=COLUMNS(P_Spill)"
			calculate
			my addLine("T06", "INFO", "references=" & (references of named item "P_Spill" of wb) & "  COLUMNS->" & ((value of range "C2" of s1) as text))
		on error e
			my addLine("T06", "FAIL", e)
		end try

		-- T08 update a definition (T07 enumerates at the end) ---------------------------
		try
			set formula2 of range "A4" of s1 to "=P_K"
			set references of named item "P_K" of wb to "=20"
			calculate
			my addLine("T08", "INFO", "after update S1!A4=" & ((value of range "A4" of s1) as text) & " (expect 20)")
		on error e
			my addLine("T08", "FAIL", e)
		end try

		-- T09 rename: do dependents follow? -------------------------------------------
		try
			set name of named item "P_K" of wb to "P_K2"
			my addLine("T09", "INFO", "renamed; S1!A4 formula now " & (formula of range "A4" of s1))
		on error e
			my addLine("T09", "FAIL", e)
		end try

		-- T10 delete --------------------------------------------------------------------
		try
			delete named item "P_K2" of wb
			calculate
			my addLine("T10", "INFO", "deleted; S1!A4 shows " & (string value of range "A4" of s1))
		on error e
			my addLine("T10", "FAIL", e)
		end try

		-- T11 syntax error: what does the caller get? ---------------------------------
		try
			make new named item at wb with properties {name:"P_Bad", references:"=LAMBDA(x, x+"}
			my addLine("T11", "INFO", "accepted (!) references=" & (references of named item "P_Bad" of wb))
		on error e
			my addLine("T11", "INFO", "rejected with: " & e)
		end try

		-- T12 collision with a built-in (E9's Fact/FACT) -------------------------------
		try
			make new named item at wb with properties {name:"Fact", references:"=LAMBDA(n, 1)"}
			set formula2 of range "A5" of s1 to "=Fact(5)"
			calculate
			my addLine("T12", "INFO", "=Fact(5) -> " & ((value of range "A5" of s1) as text) & " (1 = name wins, 120 = built-in wins); formula reads " & (formula of range "A5" of s1))
		on error e
			my addLine("T12", "INFO", "rejected: " & e)
		end try

		-- T13 dotted (AFE-module-style) name --------------------------------------------
		try
			make new named item at wb with properties {name:"Mod.Fn", references:"=LAMBDA(x, x*10)"}
			set formula2 of range "A6" of s1 to "=Mod.Fn(2)"
			calculate
			my addLine("T13", "INFO", "=Mod.Fn(2) -> " & ((value of range "A6" of s1) as text))
		on error e
			my addLine("T13", "FAIL", e)
		end try

		-- T14 line breaks inside a definition ------------------------------------------
		try
			make new named item at wb with properties {name:"P_Multi", references:"=LAMBDA(x," & nl & "  x*2)"}
			set r to references of named item "P_Multi" of wb
			set formula2 of range "A7" of s1 to "=P_Multi(4)"
			calculate
			if r contains nl then
				set keep to "newline KEPT"
			else
				set keep to "newline LOST"
			end if
			my addLine("T14", "INFO", keep & "; =P_Multi(4) -> " & ((value of range "A7" of s1) as text))
		on error e
			my addLine("T14", "FAIL", e)
		end try

		-- T15 length limit ----------------------------------------------------------------
		repeat with n in {3990, 4090, 4500}
			set f to "=LAMBDA(x,x" & my repeatText("+1", n) & ")"
			set nm to "P_Long" & (n as text)
			try
				make new named item at wb with properties {name:nm, references:f}
				set formula2 of range "A9" of s1 to "=" & nm & "(0)"
				calculate
				my addLine("T15", "INFO", (length of f) & " chars accepted; ->" & ((value of range "A9" of s1) as text))
			on error e
				my addLine("T15", "INFO", (length of f) & " chars rejected: " & e)
			end try
		end repeat

		-- T16 evaluate an arbitrary expression without touching a cell ------------------
		try
			my addLine("T16", "INFO", "evaluate P_Add1(1) -> " & ((evaluate name "P_Add1(1)") as text))
		on error e
			my addLine("T16", "FAIL", "evaluate P_Add1(1): " & e)
		end try
		try
			my addLine("T16", "INFO", "evaluate ROWS*10+COLUMNS of P_Spill -> " & ((evaluate name "ROWS(P_Spill)*10+COLUMNS(P_Spill)") as text))
		on error e
			my addLine("T16", "FAIL", "evaluate shape: " & e)
		end try

		-- T17 non-ASCII name (λ) ------------------------------------------------------
		try
			make new named item at wb with properties {name:"Grow" & lam, references:"=LAMBDA(b, g, b*(1+g))"}
			set formula2 of range "A10" of s1 to "=Grow" & lam & "(100, 0.1)"
			calculate
			my addLine("T17", "INFO", "=Growλ(100,0.1) -> " & ((value of range "A10" of s1) as text))
		on error e
			my addLine("T17", "FAIL", e)
		end try

		-- T18 write in the LOCAL language (list separator ';' in it-IT) ----------------
		try
			make new named item at wb with properties {name:"P_Loc", reference local:"=LAMBDA(x; x+1)"}
			my addLine("T18", "INFO", "reference local with ';' accepted; references=" & (references of named item "P_Loc" of wb))
		on error e
			my addLine("T18", "INFO", "reference local with ';' rejected: " & e)
		end try

		-- T07 enumerate everything ------------------------------------------------------
		try
			set c to count of named items of wb
			my addLine("T07", "INFO", (c as text) & " named items in workbook collection")
			repeat with i from 1 to c
				set ni to named item i of wb
				set rf to references of ni
				if (length of rf) > 80 then set rf to (text 1 thru 80 of rf) & "…"
				my addLine("T07", "ITEM", (name of ni) & " | visible=" & ((visible of ni) as text) & " | " & rf)
			end repeat
		on error e
			my addLine("T07", "FAIL", e)
		end try

		-- T19 save a copy for the file-level comparison, close unsaved -------------------
		try
			-- POSIX path, no file format (HFS path / explicit format: "Errore nei parametri").
			-- After save-as the workbook is renamed, so `wb` is stale: close it by its new name.
			save workbook as wb filename (saveDir & "/probe_mac.xlsx")
			my addLine("T19", "PASS", "saved " & saveDir & "/probe_mac.xlsx")
			close workbook "probe_mac.xlsx" saving no
		on error e
			my addLine("T19", "FAIL", e)
			try
				close wb saving no
			end try
		end try
	end tell
	return rpt
end run
