   'demo.vbs
   Option Explicit
   dim MyConn
   dim MdbFilePath
   dim SQL_query 
   dim UPD_query 
   dim RS
   dim savekey
   dim saveruntme
   dim compruntme
   dim compkey
   dim runplace
   dim runplaceincr
Const adUseClient = 3
   wscript.echo "Total Jobs Each User - Sorted by Largest Total First"

   Set MyConn = CreateObject("ADODB.Connection")
   MdbFilePath = "raceresults.mdb"
   MyConn.Open "Driver={Microsoft Access Driver (*.mdb)}; DBQ=" & MdbFilePath & ";"
   
   UPD_query = "UPDATE [Race-Import] SET runadjplace = 1000"
   wscript.echo UPD_query
   MyConn.Execute UPD_query

   UPD_query = "UPDATE [Race-Import] SET wcpoints = 0"
   wscript.echo UPD_query
   MyConn.Execute UPD_query

   UPD_query = "UPDATE [Race-Import] SET Gender= 'Girls' WHERE Gender='womens result export'"
   wscript.echo UPD_query
   MyConn.Execute UPD_query

   UPD_query = "UPDATE [Race-Import] SET Gender= 'Boys' WHERE Gender='mens result export'"
   wscript.echo UPD_query
   MyConn.Execute UPD_query

   SQL_query = "SELECT * FROM [Q-sorted-places]"
   Wscript.echo  sql_query
   Set RS = MyConn.Execute(SQL_query)
   
   savekey = ""
   saveruntme = "saverun"
   runplaceincr = 1
   WHILE NOT RS.EOF
   	compkey = RS("race") & RS("Gender") & RS("Class") & RS("run")
        wscript.echo "test-" & RS("runtme")
	if compkey <> savekey then
		savekey =compkey
		runplace = 0
		runplaceincr = 1
	end if
        if (saveruntme = RS("runtme")) then
		wscript.echo "sametime"
		runplaceincr = runplaceincr + 1
	else
		runplace = runplace + runplaceincr 
                runplaceincr = 1
		saveruntme = RS("runtme")
	end if
	if (RS("runplace") = 99999) then
		runplace = 1000
	end if
	wscript.echo "here 1"
	UPD_query = "UPDATE [Race-Import] SET runadjplace = " & runplace   & " WHERE ussa = '" & RS("ussa") & "' and race = '" & RS("race") & "' and run = '" & RS("run") & "'" 
	wscript.echo UPD_query
	MyConn.Execute UPD_query
        Wscript.echo "User" & RS("ussa") & "-" & RS("name") & "-" & runplace
   RS.MoveNext
   WEND

   RS.Close
   set RS = nothing  
	
rem   UPD_query = "UPDATE [Race-Import] INNER JOIN [WC-Points] ON [Race-Import].runadjplace = [WC-Points].Place SET [Race-Import].wcpoints = [WC-Points]![Points] WHERE ((([Race-Import].run)='Combined'));"
   UPD_query = "UPDATE [Race-Import] INNER JOIN [WC-Points] ON [Race-Import].runadjplace = [WC-Points].Place SET [Race-Import].wcpoints = [WC-Points]![Points] ;"
   wscript.echo UPD_query
   MyConn.Execute UPD_query

   MyConn.close
   set MyConn = nothing

