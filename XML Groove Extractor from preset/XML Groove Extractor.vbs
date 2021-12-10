' https://github.com/VoilierBleu/BFD : 11:50 vendredi 10 décembre 2021
Set o_FSO = CreateObject("Scripting.FileSystemObject")
Set o_file_in = o_FSO.GetFile(WScript.Arguments.item(0))
Set ts = o_file_in.OpenAsTextStream(1,-2)

Set xmlDoc = CreateObject("Microsoft.XMLDOM")
xmlDoc.Async = "false"
xmlDoc.Load(WScript.Arguments.item(0))
Set colNodes=xmlDoc.selectNodes ("//root/BFD2GrooveBundle/BFD2GrooveBundleInfo/@info_name")
For Each objNode in colNodes
	s_file=objNode.Text
Next
i_counter_file=0

b_file_name_found=False
s_ext=".bfd3pal"
Do While b_file_name_found=false
	If o_FSO.FileExists("D:\AUDIO\Librairies\Drums\BFD3\"&s_file&s_ext) Then
		i_counter_file=i_counter_file+1
		s_ext="("&i_counter_file&").bfd3pal"
	Else
		Set o_file_out = o_FSO.CreateTextFile("D:\AUDIO\Librairies\Drums\BFD3\"&s_file&s_ext,True)
'		WScript.Echo "D:\AUDIO\Librairies\Drums\BFD3\"&s_file&s_ext
		b_file_name_found=true		
	End If
'	If i_counter_file>100 Then wscript.Echo "KO1" : WScript.Quit
Loop

Set colNodes=xmlDoc.selectNodes ("//root/@loop")
For Each objNode in colNodes
	s_loop=objNode.Text
Next
Set colNodes=xmlDoc.selectNodes ("//root/@ppqn")
For Each objNode in colNodes
	s_ppqn=objNode.Text
Next
Set colNodes=xmlDoc.selectNodes ("//root/@timeformat")
For Each objNode in colNodes
	s_timeformat=objNode.Text
Next
Set colNodes=xmlDoc.selectNodes ("//root/@loopstart")
For Each objNode in colNodes
	s_loopstart=objNode.Text
Next
Set colNodes=xmlDoc.selectNodes ("//root/@loopend")
For Each objNode in colNodes
	s_loopend=objNode.Text
Next
Set colNodes=xmlDoc.selectNodes ("//root/@timesig")
For Each objNode in colNodes
	s_timesig=objNode.Text
Next
Set colNodes=xmlDoc.selectNodes ("//root/@bpm")
For Each objNode in colNodes
	s_bpm=objNode.Text
Next
o_file_out.WriteLine("<root timeformat="""&s_timeformat&""" ppqn="""&s_ppqn&""" loop="""&s_loop&""" loopstart="""&s_loopstart&""" loopend="""&s_loopend&""" timesig="""&s_timesig&""" bpm="""&s_bpm&""">")

b_extract_flag=false
i_counter=0
Do Until ts.AtEndOfStream
    s_line = ts.ReadLine
'    wscript.echo s_line
    i_counter=i_counter+1
'    if i_counter=1 then
'        o_file_out.WriteLine(s_line)
'    end if
    If InStr(s_line, "BFD2GrooveBundle ") > 0 Then
        b_extract_flag=true
    elseif InStr(s_line, "<BFD2PatternManagerSettings ") > 0 Then
        o_file_out.WriteLine(s_line)
        b_extract_flag=false
    End If
    If b_extract_flag=true Then
        o_file_out.WriteLine(s_line)
    End If
Loop
o_file_out.WriteLine("</root>")