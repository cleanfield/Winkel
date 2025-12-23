Dim fld, Err, st, fileSys, fileImg, fileTxt, header, answer

Const prodFile = "C:\Documents and Settings\Administrator\Desktop\stap2\xcart_prod.csv"
Const imgFile = "C:\Documents and Settings\Administrator\Desktop\stap2\xcart_img.csv"
Const imgDir = "C:\Documents and Settings\Administrator\Desktop\stap2\pict\pict_"

Const ForWriting = 2

local_int = Wscript.Arguments(0)

answer=MsgBox("Artikelen exporteren uit kopie Exact ...")

'Create a new connection object
Set Cnxn = CreateObject("ADODB.Connection")
strCnxn = "Provider='sqloledb';Data Source=localhost;Integrated Security='SSPI';Initial Catalog='WerkaandeWinkel';"

'Specify connection information
Cnxn.CommandTimeout = 9600
Cnxn.Open strCnxn


'Execute the command and retrieve the returned recordset
set rsItems = CreateObject("ADODB.recordset")

strSQLItems = "Select * from ruben_prod Where picture is not Null order by art_id"
rsItems.Open strSQLItems, Cnxn

'Initialize the stream object used to persist to a file
Set st = createobject("Adodb.Stream")
st.Type = 1

do until rsItems.EOF
    st.Open
'Write the value of the field to the stream

    st.Write rsItems.fields("picture").Value
'Save the content of the stream to a file
    st.SaveToFile imgDir & rsItems.fields("art_id").Value & ".jpg", 2

'Close the stream
    st.close
    rsItems.MoveNext
loop

'Close recordset and connection
rsItems.Close