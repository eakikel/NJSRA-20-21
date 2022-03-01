   'demo.vbs
   Option Explicit
   dim MyConn
   dim MdbFilePath
   dim SQL_query 
   dim UPD_query 
   dim DEL_query 
   dim RS
   dim savekey
   dim saveruntme
   dim compruntme
   dim compkey
   dim runplace
   dim runplaceincr
Const adUseClient = 3

   Set MyConn = CreateObject("ADODB.Connection")
   MdbFilePath = "raceresults.mdb"
   MyConn.Open "Driver={Microsoft Access Driver (*.mdb)}; DBQ=" & MdbFilePath & ";"
   
   SQL_query = 	"SELECT * FROM [Q-RacersToDelete];"_



   Wscript.echo  sql_query
   Set RS = MyConn.Execute(SQL_query)
   

   WHILE NOT RS.EOF
    wscript.echo "test-" & RS("ussa") 
	wscript.echo "here 1"
	UPD_query = "UPDATE [Race-Import] SET outofstate = True WHERE ussa = '" & RS("ussa") & "';"
	wscript.echo UPD_query
	MyConn.Execute UPD_query

   RS.MoveNext
   WEND

   RS.Close
   set RS = nothing  
	

   MyConn.close
   set MyConn = nothing

