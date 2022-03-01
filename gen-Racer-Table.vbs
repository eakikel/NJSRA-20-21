   'demo.vbs
   Option Explicit
   dim MyConn
   dim MdbFilePath
   dim SQL_query 
   dim del_query 
   dim ins_query 
   dim upd_query 
   dim RS
   dim RS_racers
   dim RS_NJSRA
Dim wrkname
Dim wrkussa
Dim wrkclub
Dim wrkclass
Dim wrkyob
Dim wrkoutofstate
Dim wrkgender
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

   wscript.echo "Generate Racer Scores"

   Set MyConn = CreateObject("ADODB.Connection")
   MdbFilePath = "raceresults.mdb"
   MyConn.Open "Driver={Microsoft Access Driver (*.mdb)}; DBQ=" & MdbFilePath & ";"

   del_query = "DELETE * FROM [Racers]"
   wscript.echo del_query
   MyConn.Execute del_query


   SQL_query = "SELECT [Race-Import].ussa, [Race-Import].club, [Race-Import].yob, [Race-Import].name, [Race-Import].class, [Race-Import].gender, [Race-Import].outofstate "
   SQL_query = SQL_query & "FROM [Race-Import] "
   SQL_query = SQL_query & "GROUP BY [Race-Import].ussa, [Race-Import].club, [Race-Import].yob, [Race-Import].name, [Race-Import].class, [Race-Import].gender, [Race-Import].outofstate; "
   Set RS = MyConn.Execute(SQL_query)
   
   
   WHILE NOT RS.EOF
	wrkname = replace(RS("name"), "'" ,"''")
	wrkussa = RS("ussa")
	wrkclass = RS("class")
	wrkclub = RS("club")
	wrkyob = RS("yob")
	wrkgender = RS("gender")
	wrkoutofstate = RS("outofstate")
	
	SQL_query = "SELECT * "
	SQL_query = SQL_query & "FROM [Q-NJSRA-Members] "
   	SQL_query = SQL_query & "Where [Q-NJSRA-Members].ussa = '" & RS("ussa") & "'"
	wscript.echo SQL_query
   	Set RS_NJSRA = MyConn.Execute(SQL_query)
	
	If RS_NJSRA.EOF Then
		wscript.echo RS("ussa") & "NOT IN NJSRA DATABASE"
		RS_NJSRA.Close
		set RS_NJSRA = Nothing
	Else
		wrkyob = RS_NJSRA("yob")
		RS_NJSRA.Close
		set RS_NJSRA = Nothing
	End IF
	

		


	
	SQL_query = "SELECT [Racers].ussa "
	SQL_query = SQL_query & "FROM [Racers] "
   	SQL_query = SQL_query & "Where [Racers].ussa = " & chr(39) & RS("ussa") & chr(39)
   	Set RS_racers = MyConn.Execute(SQL_query)

	If RS_racers.EOF Then
		ins_query = "INSERT INTO [Racers](ussa,name,class,club,yob,Gender,outofstate) "
		ins_query = ins_query & "values (" 
		ins_query = ins_query & chr(39) & RS("ussa") & chr(39)& "," 
		ins_query = ins_query & chr(39) & wrkname & chr(39)& "," 
		ins_query = ins_query & chr(39) & RS("class") & chr(39)& "," 
		ins_query = ins_query & chr(39) & RS("club") & chr(39)& "," 
		ins_query = ins_query & chr(39) & wrkyob & chr(39)& "," 
		ins_query = ins_query & chr(39) & RS("gender") & chr(39)& "," 
		ins_query = ins_query & RS("outofstate") & ")" 
   		wscript.echo ins_query
   		MyConn.Execute ins_query
		RS_racers.Close
		set RS_racers = nothing  

    	 	wscript.echo RS("ussa") & " does not exist"
  	Else
		RS_racers.Close
		set RS_racers = nothing  
		upd_query = "UPDATE [Racers] "
		upd_query = upd_query & "SET "
		upd_query = upd_query & " name = " & chr(39) & wrkname & chr(39)& "," 
		upd_query = upd_query & " class = " & chr(39) & RS("class") & chr(39)& "," 
		upd_query = upd_query & " club = " & chr(39) & RS("club") & chr(39)& "," 
		upd_query = upd_query & " yob = " & chr(39) & wrkyob & chr(39)& "," 
		upd_query = upd_query & " gender = " & chr(39) & RS("gender") & chr(39)
		upd_query = upd_query & " WHERE ussa = " & chr(39) & wrkussa & chr(39) 
		wscript.echo upd_query
		MyConn.Execute upd_query

    	 	wscript.echo RS("ussa") & " exists and updates"
  	End If
   RS.MoveNext
   WEND

   RS.Close
   set RS = nothing  
   MyConn.close
   set MyConn = nothing

