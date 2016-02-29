function GetFso() {
    var result = window.fso
    if (result == null) {
        window.fso = new ActiveXObject("Scripting.FileSystemObject")
        result = window.fso
    }
    return result;
}
function GetFullPath(folderPath){
    var fso = GetFso()
    var folder = fso.GetFolder(folderPath)
    return folder.Path
}
function GetFiles(folderPath) {
    var result = {}
    result.found = [];
    result.error = "";

    var fso = GetFso()

    if (!fso.FolderExists(folderPath)) {
        result.error = "Directory does not exist."
        return result;
    }
    var folder = fso.GetFolder(folderPath)
    var fileCollection = folder.Files
    if (fileCollection.Count < 1) {
        result.error = "Directory empty."
        return result;
    }

    var fileColEnum = new Enumerator(fileCollection);
    for (var ext; !fileColEnum.atEnd() ; fileColEnum.moveNext()) {
        ext = fileColEnum.item().Name
        if (ext.length < 4)
            continue;
        ext = ext.substr(ext.length - 4)
        if (ext == ".csv")
            result.found.push(fileColEnum.item().Path)
    }
    return result;
}
function DeleteFile(filePath) {
    GetFso().DeleteFile(filePath)
}
function MoveFile(oldPath, newPath) {
    GetFso().MoveFile(oldPath, newPath)
}
function CreateNewTextFile(filePath) {
    var fso = GetFso()
    return fso.CreateTextFile(filePath, true)
}
function OpenTextFile_ReadOnly(filePath) {
    var ForReading = 1
    var fso = GetFso()
    return fso.OpenTextFile(filePath, ForReading)
}
function SplitCsvLine(line) {
    var result = []
    var index_comma
    var index_quote
    var item = ""
    while (true) {
        index_comma = line.indexOf(",")
        index_quote = line.indexOf('"')
        if (index_comma < 0) {
            //no more commas
            item += line
            result.push(item)
            return result
        }
        if (index_quote < 0) {
            //no more quotes
            item += line.substr(0, index_comma)
            line = line.substr(index_comma + 1)
            result.push(item)
            var subResult = line.split(",")
            for(var i = 0; i < subResult.length; i++)
                result.push(subResult[i])
            return result
        }
        if (index_comma < index_quote) {
            //item ends before next quote
            item += line.substr(0, index_comma)
            line = line.substr(index_comma+1)
            result.push(item)
            item = ""
        } else {
            //quote starts escape
            item += line.substr(0, index_quote + 1)
            line = line.substr(index_quote + 1)
            index_quote = line.indexOf('"')
            if (index_quote < 0) {
                // line ends before unescape
                item += line
                result.push(item)
                return result
            } else {
                item += line.substr(0, index_quote + 1)
                line = line.substr(index_quote + 1)
            }
        }

    }
}