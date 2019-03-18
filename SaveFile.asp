<%
dim file,filename,houzui
file = 
Request.Form("file")
 
if file="" 
then
response.write"<script>alert('Choose the file you wanna upload');window.location.href='upload.htm';</script>"
else
houzui=mid(file,InStrRev(file, 
"."))

filename=year(date) & month(date) & day(date) & 
Hour(time) & minute(time) & second(time) & houzui
 
Set 
objStream = Server.CreateObject("ADODB.Stream")
objStream.Type = 
1
objStream.Open
objStream.LoadFromFile file
objStream.SaveToFile 
Server.MapPath(filename),2
objStream.Close
 
 
response.write"<script>alert('upload success');window.location.href='upload.htm';</script>"
else
response.write"<script>alert('upload failed');window.location.href='upload.htm';</script>"
end if
end 
if
%>
