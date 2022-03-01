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
   dim score_class
   dim tie_score_class
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
   dim wrkscore(10)

   minruns = int(WScript.Arguments.Item(0))
   wscript.echo "Generate Racer Scores" & "-" & minruns 

   Set MyConn = CreateObject("ADODB.Connection")
   MdbFilePath = "raceresults.mdb"
   MyConn.Open "Driver={Microsoft Access Driver (*.mdb)}; DBQ=" & MdbFilePath & ";"

   del_query = "DELETE * FROM [Racer-Score-U12]"
   wscript.echo del_query
   MyConn.Execute del_query

   SQL_query = "SELECT * FROM [racer-runs-U12]"
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
		wrkscore(1) = 999
		wrkscore(2) = 999
		wrkscore(3) = 999
		wrkscore(4) = 999
		wrkscore(5) = 999
		wrkscore(6) = 999
		wrkscore(7) = 999
		wrkscore(8) = 999
		wrkscore(9) = 999
		wrkscore(10) = 999

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
		wrkscore(1) = 999
		wrkscore(2) = 999
		wrkscore(3) = 999
		wrkscore(4) = 999
		wrkscore(5) = 999
		wrkscore(6) = 999
		wrkscore(7) = 999
		wrkscore(8) = 999
		wrkscore(9) = 999
		wrkscore(10) = 999
	end if
        
       	
	Select Case RS("runplace") < 1000
		Case true
			runcount = runcount+1
			wscript.echo RS("ussa") & "-" & runcount & "-" & RS("runplace")
			Select Case runcount > minruns
				Case false
   					score_class = score_class + RS("runadjplace") 
				Case true
  					wrkscore(runcount - minruns) = + RS("runadjplace") 
				end select
		Case false
			if runcount < minruns then
				score_class = score_class + RS("runadjplace") 
			end if
	end select

 	wscript.echo RS("ussa") & "-" & score_class & "-" & runcount
   RS.MoveNext
   WEND
   call writescore
   RS.Close
   set RS = nothing  
   MyConn.close
   set MyConn = nothing	

sub writescore
	If runcount < minruns then
		score_class = 999
	end if
	wscript.echo "wrkScore1-" & wrkscore(1) & "-wrkScore2-" & wrkscore(2)
        tie_score_class  = wrkscore(1) /1000 + wrkscore(2) /1000000 + wrkscore(3) /1000000000 
	testname = Replace(savename,"'","''")
	wscript.echo "testname-" & testname
	ins_query = "INSERT INTO [Racer-Score-U12](ussa,name,class,gender,score_class,tie_score_class) "
	ins_query = ins_query & "values ('" & saveussa & "','" & testname & "',"
	ins_query = ins_query & "'" & saveclass & "',"
	ins_query = ins_query & "'" & savegender & "',"
	ins_query = ins_query & "'" & score_class & "',"
	ins_query = ins_query & "'" & tie_score_class  & "')"

   	wscript.echo ins_query
   	MyConn.Execute ins_query
	wscript.echo "Score-" & saveussa & "-" & score_class & "-" & runcount
end sub









