Class Preset
    Public Name
    Public Directory
    Public SummaryFilename
    Public FtpHost
    Public FtpDirectory
    Public FtpUsername
    Public FtpPassword
End Class

Set Presets = CreateObject("Scripting.Dictionary")
Sub LoadPresets
    'Set window.Presets = CreateObject("Scripting.Dictionary")
    Set xDoc = CreateObject( "Microsoft.XMLDOM" )
	xDoc.async = False
	xDoc.validateOnParse = False
	If xDoc.Load("src/presets.xml") Then
	    ' The document loaded successfully.
	    Set xmlPresets = xDoc.getElementsByTagName("preset")
        Set objPreset = Nothing
        For Each xmlPreset In xmlPresets
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
End Sub

Sub AddPresetOption(byval text, byval value)
    Set drp = document.getElementById("optPreset")
    If drp Is Nothing Then Exit Sub

    Set opt = document.createElement("option")
    opt.value = value
    opt.text = text
    drp.options.add opt
End Sub

Sub optPreset_Change
    Set drp = document.getElementById("optPreset")
    If drp Is Nothing Then Exit Sub
    Set objPreset = window.Presets.Item(drp.value)
    If objPreset Is Nothing Then 
        Set out = document.getElementById("ErrorOut")
        If out Is Nothing Then 
            alert "Could not find preset."
        Else
            out.innerHTML = "Could not find preset."
        End IF
        drp.selectedIndex = 0
        Exit Sub
    end If

    Set el = document.getElementById("txtDir")
    If Not el Is Nothing Then el.value = objPreset.Directory
    Set el = document.getElementById("txtSummaryFileName")
    If Not el Is Nothing Then el.value = objPreset.SummaryFilename
    Set el = document.getElementById("txtFtpHost")
    If Not el Is Nothing Then el.value = objPreset.FtpHost
    Set el = document.getElementById("txtFtpDirectory")
    If Not el Is Nothing Then el.value = objPreset.FtpDirectory
    Set el = document.getElementById("txtFtpUser")
    If Not el Is Nothing Then el.value = objPreset.FtpUsername
    Set el = document.getElementById("txtFtpPassword")
    If Not el Is Nothing Then el.value = objPreset.FtpPassword

    txtDir_Change
    txtSummaryFileName_Change
End Sub
