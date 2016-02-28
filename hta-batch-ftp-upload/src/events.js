function fileSelector_Change(fileSelectorName, txtName) {
    var txt = document.getElementById(txtName);
    if (txt == null)
        return;
    var filePicker = document.getElementById(fileSelectorName);
    if (filePicker == null)
        return;
    var val = filePicker.value;
    val = val.substr(0, val.lastIndexOf('\\'));
    txt.value = val;
    txt.onchange();
}
function txtSummaryFileName_Change(txtFromName, txtToName) {
    var txtFrom = document.getElementById(txtFromName);
    if (txtFrom == null)
        return;
    var txtTo = document.getElementById(txtToName);
    if (txtTo == null)
        return;
    var now = new Date();
    var year = now.getFullYear();
    var month = now.getMonth() + 1 + "";
    if (month.length < 2)
        month = "0" + month;
    var day = now.getDate() + "";
    if (day.length < 2)
        day = "0" + day;
    var hour = now.getHours() + "";
    if (hour.length < 2)
        hour = "0" + hour;
    var minute = now.getMinutes() + "";
    if (minute.length < 2)
        minute = "0" + minute;
    var second = now.getSeconds() + "";
    if (second.length < 2)
        second = "0" + second;
    txtTo.value = txtFrom.value.replace(/{YEAR}/gi, year)
                               .replace(/{MONTH}/gi, month)
                               .replace(/{DAY}/gi, day)
                               .replace(/{HOUR}/gi, hour)
                               .replace(/{MINUTE}/gi, minute)
                               .replace(/{SECONDS?}/gi, second)
                               .replace(/[^a-z0-9\-_'(),.]/gi, "");
}
function btnPreset_Click() {
    var frm = document.getElementById("frmMain");
    if (frm != null)
        frm.style.display = "none";
    frm = document.getElementById("frmPresets");
    if (frm != null)
        frm.style.display = "";
    document.body.style.backgroundColor = "#7FFFD4"; //aquamarine
}
//function btnNotPreset_Click() {
//    var frm = document.getElementById("frmMain");
//    if (frm != null)
//        frm.style.display = "";
//    frm = document.getElementById("frmPresets");
//    if (frm != null)
//        frm.style.display = "none";
//    document.body.style.backgroundColor = "";
//}
function lstPresets_Change() {
    var lst = document.getElementById("lstPresets")
    if (lst == null)
        return
    //var objPreset = window.Presets_js[lst.value]
    var objPreset = window.Presets.Item(lst.value)
    if (objPreset == null) {
        var out = document.getElementById("ErrorOut")
        if (out == null)
            alert("Could not find preset.")
        else
            out.innerHTML = "Could not find preset."
        //objPreset = window.Presets_js["None"]
        objPreset = window.Presets.Item("None");
    }

    var el = document.getElementById("txtEditName")
    if (el != null) {
        if (objPreset.Name == "New")
            el.defaultValue = ""
        else
            el.defaultValue = objPreset.Name
    }
    el.value = el.defaultValue;
    var el = document.getElementById("txtEditDir")
    if (el != null)
        el.value = objPreset.Directory
    el = document.getElementById("txtEditSummaryFileName")
    if (el != null)
        el.value = objPreset.SummaryFilename
    el = document.getElementById("txtEditFtpHost")
    if (el != null)
        el.value = objPreset.FtpHost
    el = document.getElementById("txtEditFtpDirectory")
    if (el != null)
        el.value = objPreset.FtpDirectory
    el = document.getElementById("txtEditFtpUser")
    if (el != null)
        el.value = objPreset.FtpUsername
    el = document.getElementById("txtEditFtpPassword")
    if (el != null)
        el.value = objPreset.FtpPassword

    txtDir_Change("txtEditDir", "lblEditDirectoryStatus")
    txtSummaryFileName_Change("txtEditSummaryFileName", "txtEditSummaryFileNamePreview")
    if (objPreset.Name == "New")
        btnEditPreset_Click()
    buttonStateChange_PresetNotEdit()
}
function buttonStateChange_PresetEdit() {
    var btnEdit = document.getElementById('btnEditPreset');
    if (btnEdit != null)
        btnEdit.disabled = true;
    var btnSave = document.getElementById('btnSaveEditPreset');
    if (btnSave != null)
        btnSave.disabled = false;
    var btnCancel = document.getElementById('btnCancelEditPreset');
    if (btnCancel != null)
        btnCancel.disabled = false;
    el = document.getElementById('txtEditName');
    if (el != null)
        el.disabled = false;
    el = document.getElementById('txtEditDir');
    if (el != null)
        el.disabled = false;
    el = document.getElementById('EditfileSelector');
    if (el != null)
        el.disabled = false;
    el = document.getElementById('txtEditSummaryFileName');
    if (el != null)
        el.disabled = false;
    el = document.getElementById('txtEditFtpHost');
    if (el != null)
        el.disabled = false;
    el = document.getElementById('txtEditFtpUser');
    if (el != null)
        el.disabled = false;
    el = document.getElementById('txtEditFtpPassword');
    if (el != null)
        el.disabled = false;
    el = document.getElementById('txtEditFtpDirectory');
    if (el != null)
        el.disabled = false;
}
function buttonStateChange_PresetNotEdit() {
    var el = document.getElementById('btnEditPreset');
    if (el != null)
        el.disabled = false;
    el = document.getElementById('btnSaveEditPreset');
    if (el != null)
        el.disabled = true;
    el = document.getElementById('btnCancelEditPreset');
    if (el != null)
        el.disabled = true;
    el = document.getElementById('txtEditName');
    if (el != null)
        el.disabled = true;
    el = document.getElementById('txtEditDir');
    if (el != null)
        el.disabled = true;
    el = document.getElementById('EditfileSelector');
    if (el != null)
        el.disabled = true;
    el = document.getElementById('txtEditSummaryFileName');
    if (el != null)
        el.disabled = true;
    el = document.getElementById('txtEditFtpHost');
    if (el != null)
        el.disabled = true;
    el = document.getElementById('txtEditFtpUser');
    if (el != null)
        el.disabled = true;
    el = document.getElementById('txtEditFtpPassword');
    if (el != null)
        el.disabled = true;
    el = document.getElementById('txtEditFtpDirectory');
    if (el != null)
        el.disabled = true;
}
function btnEditPreset_Click() {
    buttonStateChange_PresetEdit()
}
function btnSaveEditPreset_Click() {
    var preset = {}
    preset.Name = trim(document.getElementById("txtEditName").value)
    if (preset.Name == "") {
        alert("you must enter a name for the preset")
        return;
    }
    preset.Directory = trim(document.getElementById("txtEditDir").value)
    preset.SummaryFilename = trim(document.getElementById("txtEditSummaryFileName").value)
    preset.FtpHost = trim(document.getElementById("txtEditFtpHost").value)
    preset.FtpDirectory = trim(document.getElementById("txtEditFtpDirectory").value)
    preset.FtpUsername = trim(document.getElementById("txtEditFtpUser").value)
    preset.FtpPassword = trim(document.getElementById("txtEditFtpPassword").value)
    if (SavePreset(preset))
        window.location.reload()
    alert("Save failed.");
    //buttonStateChange_PresetNotEdit()
}
function trim(str) {
    return str.replace(/^\s+/g, "").replace(/\s+$/g, "")
}
function txtEditName_Change() {
    // check preset does not already exist with same name
    var lblResult = document.getElementById('lblEditNameValidation');
    lblResult.innerHTML = "&nbsp;";
    var txt = document.getElementById('txtEditName')
    if (txt == null)
        return;
    txt.value = txt.value.replace(/[^a-z0-9\-_'(),. ]/gi, "")
    var lst = document.getElementById('lstPresets')
    if (lst == null)
        return;

    var name = txt.value.toLowerCase()
    if (name == "new") {
        lblResult.innerHTML = 'You may not use the name "New"'
        txt.value = txt.defaultValue
        return;
    }
    if (name == "none") {
        lblResult.innerHTML = 'You may not use the name "None"'
        txt.value = txt.defaultValue
        return;
    }
    if (lst.value.toLowerCase() == name)
        return;

    var options = lst.options;
    for (var i = 0; i < options.length; i++) {
        if (options[i].value.toLowerCase() == name) {
            lblResult.innerHTML = 'The name "' + options[i].value + '" already exists'
            txt.value = txt.defaultValue
            return;
        }
    }
}
function optPreset_Change() {
    var drp = document.getElementById("optPreset")
    if (drp == null)
        return
    //var objPreset = window.Presets_js[drp.value]
    var objPreset = window.Presets.Item(drp.value)
    if (objPreset == null) {
        var out = document.getElementById("ErrorOut")
        if (out == null)
            alert("Could not find preset.")
        else
            out.innerHTML = "Could not find preset."
        drp.selectedIndex = 0
        return
    }

    var el = document.getElementById("txtDir")
    if (el != null)
        el.value = objPreset.Directory
    el = document.getElementById("txtSummaryFileName")
    if (el != null)
        el.value = objPreset.SummaryFilename
    el = document.getElementById("txtFtpHost")
    if (el != null)
        el.value = objPreset.FtpHost
    el = document.getElementById("txtFtpDirectory")
    if (el != null)
        el.value = objPreset.FtpDirectory
    el = document.getElementById("txtFtpUser")
    if (el != null)
        el.value = objPreset.FtpUsername
    el = document.getElementById("txtFtpPassword")
    if (el != null)
        el.value = objPreset.FtpPassword

    window.setTimeout(function () { txtDir_Change("txtDir", "lblDirectoryStatus") }, 0)
    window.setTimeout(function () { txtSummaryFileName_Change("txtSummaryFileName", "txtSummaryFileName") }, 0)
}
function txtDir_Change( txtName, lblName ){
    var txt = document.getElementById(txtName)
    if (txt == null) 
        return
    var lbl = document.getElementById(lblName)
    if (lbl == null) 
        return

    if (txt.value == "") {
        lbl.innerHTML = "Unknown"
        return
    }

    var fso = window.fso
    if (fso == null) {
        window.fso = new ActiveXObject("Scripting.FileSystemObject")
        fso = window.fso
    }
    if (!fso.FolderExists(txt.value)){
        lbl.innerHTML = "Directory Not Found"
        return
    }
                
    var folder = fso.GetFolder(txt.value)
    var files = folder.Files
    if (files.Count < 1){
        lbl.innerHTML = "Directory Empty"
        return
    }

    var count = 0
    var ext = ""
    var en = new Enumerator(files);
    for (; !en.atEnd() ; en.moveNext()) {
        ext = en.item().Name
        if (ext.length < 4)
            continue;
        ext = ext.substr(ext.length - 4)
        if (ext == ".csv")
            count++
    }
    lbl.innerHTML = count + " files found"
}