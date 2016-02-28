Class Preset
    Public Name
    Public Directory
    Public SummaryFilename
    Public FtpHost
    Public FtpDirectory
    Public FtpUsername
    Public FtpPassword
End Class

Public Presets
Set Presets = CreateObject("Scripting.Dictionary")
Sub LoadPresets
    'AddPreset_js false, "None", "", "", "", "", "", ""
    Set objPreset = New Preset
    objPreset.Name = "None"
    objPreset.Directory = ""
    objPreset.SummaryFilename = ""
    objPreset.FtpHost = ""
    objPreset.FtpDirectory = ""
    objPreset.FtpUsername = ""
    objPreset.FtpPassword = ""
    window.Presets.Add objPreset.Name, objPreset
    'AddPreset_js false, "New", "", "", "", "", "", ""
    Set objPreset = New Preset
    objPreset.Name = "New"
    objPreset.Directory = ""
    objPreset.SummaryFilename = ""
    objPreset.FtpHost = ""
    objPreset.FtpDirectory = ""
    objPreset.FtpUsername = ""
    objPreset.FtpPassword = ""
    window.Presets.Add objPreset.Name, objPreset

    Set xDoc = CreateObject( "MSXML2.DOMDocument" )
	xDoc.async = False
	xDoc.validateOnParse = False
	If xDoc.Load("src/presets.xml") Then
	    ' The document loaded successfully.
	    Set xmlPresets = xDoc.getElementsByTagName("preset")
        For Each xmlPreset In xmlPresets
'            AddPreset_js true, xmlPreset.getAttribute("name"), _ 
'                            xmlPreset.getElementsByTagName("Directory").Item(0).Text,  _
'                            xmlPreset.getElementsByTagName("SummaryFilename").Item(0).Text,  _
'                            xmlPreset.getElementsByTagName("FtpHost").Item(0).Text,  _
'                            xmlPreset.getElementsByTagName("FtpDirectory").Item(0).Text,  _
'                            xmlPreset.getElementsByTagName("FtpUsername").Item(0).Text,  _
'                            xmlPreset.getElementsByTagName("FtpPassword").Item(0).Text
            Set objPreset = New Preset
            objPreset.Name = xmlPreset.getAttribute("name")
            objPreset.Directory = xmlPreset.getElementsByTagName("Directory").Item(0).Text
            objPreset.SummaryFilename = xmlPreset.getElementsByTagName("SummaryFilename").Item(0).Text
            objPreset.FtpHost = xmlPreset.getElementsByTagName("FtpHost").Item(0).Text
            objPreset.FtpDirectory = xmlPreset.getElementsByTagName("FtpDirectory").Item(0).Text
            objPreset.FtpUsername = xmlPreset.getElementsByTagName("FtpUsername").Item(0).Text
            objPreset.FtpPassword = xmlPreset.getElementsByTagName("FtpPassword").Item(0).Text
            
            window.Presets.Add objPreset.Name, objPreset
            AddPresetOption objPreset.Name, objPreset.Name
        Next
	Else
		' The document failed to load.
        Set out = document.getElementById("ErrorOut")
        If Not out Is Nothing Then 
		    Set xPE = xDoc.parseError
		    With xPE
			    errorOut.innerHTML = "Failed to load presets " & _
				    "due the following error." & "<br />" & _
				    "Error #: " & .errorCode & ": " & xPE.reason & _
				    "Line #: " & .Line & "<br />" & _
				    "Line Position: " & .linepos & "<br />" & _
				    "Position In File: " & .filepos & "<br />" & _
				    "Source Text: " & .srcText & "<br />" & _
				    "Document URL: " & .url
		    End With
		    Set xPE = Nothing
        End If
	End If
	Set xDoc = Nothing
    optPreset_Change
    lstPresets_Change
End Sub

Sub AddPresetOption(byval text, byval value)
    Set drp = document.getElementById("optPreset")
    If drp Is Nothing Then Exit Sub
    Set lst = document.getElementById("lstPresets")
    If lst Is Nothing Then Exit Sub
    
    Set opt = document.createElement("option")
    opt.value = value
    opt.text = text
    drp.options.add opt
    Set opt = document.createElement("option")
    opt.value = value
    opt.text = text
    lst.options.add opt, lst.options.length - 1
End Sub

Function SavePreset( objPreset )
    Set xDoc = CreateObject( "MSXML2.DOMDocument" )
	xDoc.async = False
	xDoc.validateOnParse = False
	If xDoc.Load("src/presets.xml") Then
	    ' The document loaded successfully.
        Set root = xDoc.getElementsByTagName("presets").Item(0)
	    Set xmlPresets = xDoc.getElementsByTagName("preset")
        lowername = LCase(objPreset.Name)
        For Each xmlPreset In xmlPresets
            If (LCase(xmlPreset.getAttribute("name")) = lowername ) Then
                root.RemoveChild xmlPreset
            End If
        Next

        Set xmlPreset = xDoc.CreateElement("preset")
        xmlPreset.SetAttribute "name", objPreset.Name
        xmlPreset.AppendChild xDoc.CreateTextNode(vbNewLine & vbTab)
        root.AppendChild xDoc.CreateTextNode(vbNewLine & vbTab)
        root.AppendChild xmlPreset
        root.AppendChild xDoc.CreateTextNode(vbNewLine)
        
        AddXmlEl xDoc, xmlPreset, "Directory", objPreset.Directory
        AddXmlEl xDoc, xmlPreset, "SummaryFilename", objPreset.SummaryFilename
        AddXmlEl xDoc, xmlPreset, "FtpHost", objPreset.FtpHost
        AddXmlEl xDoc, xmlPreset, "FtpDirectory", objPreset.FtpDirectory
        AddXmlEl xDoc, xmlPreset, "FtpUsername", objPreset.FtpUsername
        AddXmlEl xDoc, xmlPreset, "FtpPassword", objPreset.FtpPassword
               
        xDoc.Save "src/presets.xml"
        SavePreset = True
	Else
		' The document failed to load.
        Set out = document.getElementById("ErrorOut")
        If Not out Is Nothing Then 
		    Set xPE = xDoc.parseError
		    With xPE
			    errorOut.innerHTML = "Failed to load presets " & _
				    "due the following error." & "<br />" & _
				    "Error #: " & .errorCode & ": " & xPE.reason & _
				    "Line #: " & .Line & "<br />" & _
				    "Line Position: " & .linepos & "<br />" & _
				    "Position In File: " & .filepos & "<br />" & _
				    "Source Text: " & .srcText & "<br />" & _
				    "Document URL: " & .url
		    End With
		    Set xPE = Nothing
        End If
        SavePreset = False
	End If
	Set xDoc = Nothing
End Function

Sub AddXmlEl( doc, xmlParent, name, text)
    Set el = doc.CreateElement(name)
    el.AppendChild doc.createTextNode(text)

    xmlParent.AppendChild doc.CreateTextNode(vbTab)
    xmlParent.AppendChild(el)
    xmlParent.AppendChild doc.CreateTextNode(vbNewLine & vbTab)
End Sub


Function DeletePreset( name )
    Set xDoc = CreateObject( "MSXML2.DOMDocument" )
	xDoc.async = False
	xDoc.validateOnParse = False
	If xDoc.Load("src/presets.xml") Then
	    ' The document loaded successfully.
        Set root = xDoc.getElementsByTagName("presets").Item(0)
	    Set xmlPresets = xDoc.getElementsByTagName("preset")
        lowername = LCase(name)
        For Each xmlPreset In xmlPresets
            If (LCase(xmlPreset.getAttribute("name")) = lowername ) Then
                root.RemoveChild xmlPreset
            End If
        Next
        xDoc.Save "src/presets.xml"
        DeletePreset = True
	Else
        DeletePreset = False
	End If
	Set xDoc = Nothing
End Function

Sub AddXmlEl( doc, xmlParent, name, text)
    Set el = doc.CreateElement(name)
    el.AppendChild doc.createTextNode(text)

    xmlParent.AppendChild doc.CreateTextNode(vbTab)
    xmlParent.AppendChild(el)
    xmlParent.AppendChild doc.CreateTextNode(vbNewLine & vbTab)
End Sub