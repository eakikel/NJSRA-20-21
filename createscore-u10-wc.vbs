   'demo.vbs
   Option Explicit
   dim MyConn
   dim MdbFilePath
   dim SQL_query 
   dim del_query 
   dim ins_query 
   dim RS
   dim savekey
   dim compkey
   dim runcount
   dim unused_runcount
   dim score1
   dim score5
   dim score6
   dim score7
   dim score8
   dim score9
   dim score10
   dim saveussa
   dim savename
   dim saveclass
   dim savegender
   dim testname
   dim minruns
   dim maxruns
   dim tie_score_class
   dim score_class
   dim wrkscore(10)
   dim sl_switch
   dim gs_switch
   dim sg_switch
   dim sl_switch_all
   dim gs_switch_all
   dim sg_switch_all
   dim event_cnt	
   dim unused_score_class


   minruns = int(WScript.Arguments.Item(0))
   wscript.echo "Generate Racer Scores" & "-" & minruns 


   Set MyConn = CreateObject("ADODB.Connection")
   MdbFilePath = "raceresults.mdb"
   MyConn.Open "Driver={Microsoft Access Driver (*.mdb)}; DBQ=" & MdbFilePath & ";"

   del_query = "DELETE * FROM [Racer-Score-U10] "
   wscript.echo del_query
   MyConn.Execute del_query

   SQL_query = "SELECT * FROM [racer-runs-u10]"
   Set RS = MyConn.Execute(SQL_query)
   
   savekey = ""
   WHILE NOT RS.EOF
   	compkey = RS("ussa")
	if savekey = "" then
		wscript.echo "Null save key"
		savekey =compkey
		saveussa = RS("ussa")
		savename = RS("name")
		saveclass = RS("class")
		savegender = RS("gender")
		score_class = 0
		runcount = 0
		wrkscore(1) = 0
		wrkscore(2) = 0
		wrkscore(3) = 0
		wrkscore(4) = 0
		wrkscore(5) = 0
		wrkscore(6) = 0
		wrkscore(7) = 0
		wrkscore(8) = 0
		wrkscore(9) = 0
		wrkscore(10) = 0
        sl_switch = 0
        gs_switch = 0
        sg_switch = 0
        sl_switch_all = 0
        gs_switch_all = 0
        sg_switch_all = 0
		unused_runcount = 0
		unused_score_class = 0


	end if

	if compkey <> savekey then
		
		call writescore
		savekey =compkey
		saveussa = RS("ussa")
		savename = RS("name")
		saveclass = RS("class")
		savegender = RS("gender")
		runcount = 0
		score_class = 0
		wrkscore(1) = 0
		wrkscore(2) = 0
		wrkscore(3) = 0
		wrkscore(4) = 0
		wrkscore(5) = 0
		wrkscore(6) = 0
		wrkscore(7) = 0
		wrkscore(8) = 0
		wrkscore(9) = 0
		wrkscore(10) = 0
        sl_switch = 0
        gs_switch = 0
        sg_switch = 0
        sl_switch_all = 0
        gs_switch_all = 0
        sg_switch_all = 0
		unused_runcount = 0
		unused_score_class = 0

	end if
        

	runcount = runcount+1
	wscript.echo RS("ussa") & "-" & runcount & "-" & RS("wcpoints")
	Select Case True
	Case runcount < minruns 
   		score_class = score_class + RS("wcpoints") 
	Case runcount = minruns
   		score_class = score_class + RS("wcpoints") 
 	Case runcount > minruns 
			unused_score_class = unused_score_class + RS("wcpoints")
			unused_runcount = unused_runcount+1
   			wrkscore(unused_runcount) = + RS("wcpoints") 
  	end select


   RS.MoveNext
   WEND
   call writescore
   RS.Close
   set RS = nothing  
   MyConn.close
   set MyConn = nothing

sub writescore
	event_cnt = sl_switch + gs_switch 
	wscript.echo saveussa & "-eventcnt-" & event_cnt
	score_class = score_class 
	If runcount < minruns then
		score_class = 0
	end if
Rem	If event_cnt < 2 then
Rem		score_class = 0
Rem	end if
    tie_score_class  = wrkscore(1) /1000 + wrkscore(2) /1000000 
	testname = Replace(savename,"'","''")
	wscript.echo "testname-" & testname
	ins_query = "INSERT INTO [Racer-Score-U10](ussa,name,class,gender,score_class,tie_score_class) "
	ins_query = ins_query & "values ('" & saveussa & "','" & testname & "',"
	ins_query = ins_query & "'" & saveclass & "',"
	ins_query = ins_query & "'" & savegender & "',"
	ins_query = ins_query & "'" & score_class & "',"
	ins_query = ins_query & "'" & tie_score_class & "')"

   	wscript.echo ins_query
   	MyConn.Execute ins_query
	wscript.echo "Score-" & saveussa & "-**" & event_cnt & "-**" & score_class  & "-" & runcount
end sub









