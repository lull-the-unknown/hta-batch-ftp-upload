function CreateSummaryFile_MergeAll_EnforceHeaders(filePath, folderPath) {
    folderPath = GetFullPath(folderPath)
    if (folderPath.substr(folderPath.length - 1) != "\\")
        folderPath += "\\"
    filePath = folderPath + filePath
    var files = GetFiles(folderPath);
    if (files.error != "") {
        alert("Could not create summary file, " + files.error.toLowerCase() + "\n\n" + folderPath)
        return false;
    }
    files = files.found;
    if (files.length < 1) {
        alert("Could not create summary file, no .csv files found.\n\n" + folderPath)
        return false;
    }

    var summary = null;
    var file = null;
    try {
        summary = CreateNewTextFile(filePath);
        var currFilePath = files.pop()
        if (currFilePath == filePath)
            currFilePath = files.pop()
        alert("opening " + currFilePath)
        file = OpenTextFile_ReadOnly(currFilePath)
        var headerLine = file.ReadLine()
        summary.WriteLine(headerLine)
        headerLine = SplitCsvLine(file.ReadLine())
        var headers = []
        for (var i = 0; i < headerLine.length; i++) {
            headerLine[i] = normalizeHeader(headerLine[i])
            headers[headerLine[i]] = i;
            alert(headerLine[i] + " = " + headers[headerLine[i]])
        }
        while( !file.AtEndOfStream )
            summary.WriteLine(file.ReadLine())
        file.Close()
        file = null
        while (files.length > 0) {
            var lineCount = 0
            currFilePath = files.pop()
            if (currFilePath == filePath)
                currFilePath = files.pop()
            alert("opening " + currFilePath)
            file = OpenTextFile_ReadOnly(currFilePath)
            headerLine = SplitCsvLine(file.ReadLine())
            lineCount++;
            var needsHeadersUpdated = false;
            for (var i = 0; i < headerLine.length; i++) {
                headerLine[i] = normalizeHeader(headerLine[i])
                if (headers[headerLine[i]] == null) {
                    headers[headerLine[i]] = -1
                    headers.push(headerLine[i])
                    needsHeadersUpdated = true
                }else
                    alert(headerLine[i] +" = "+headers[headerLine[i]])
            }
            if (needsHeadersUpdated) {
                var result = UpdateHeaders(summary, headers, filePath)
                summary = result.summary
                headers = result.headers
            }

            while (!file.AtEndOfStream) {
                var dataToWrite = []
                for (var i = 0; i < headers.length; i++) {
                    dataToWrite.push("")
                }
                var dataRead = SplitCsvLine(file.ReadLine())
                lineCount++;
                if (dataRead.length < headerLine.length) {
                    alert("Could not create summary file, an error occurred while processing one of the data files:\n\n"+
                            "number of data items(" + dataRead.length + ") on line " + lineCount + " does not match number " +
                            "of items(" + header.length + ") in header row.\n\n" + currFilePath)
                    return false;
                }
                for (var i = 0; i < dataRead.length; i++) {
                    var headerForCurrItem = headerLine[i]
                    var dataIndexForHeader = headers[headerForCurrItem] 
                    dataToWrite[dataIndexForHeader] = dataRead[i]
                }
                summary.WriteLine( dataToWrite.join(',') )
            }

            file.Close()
            file = null
        }
        summary.Close()
        return true
    } finally {
        if (summary != null)
            summary.Close()
        if (file != null)
            file.Close()
    }
}
function UpdateHeaders(summary, headers, summaryFilePath) {    
    var tempFile
    try {
        summary.Close()
        summary = null
        var tempFilePath = summaryFilePath + ".temp"
        MoveFile(summaryFilePath, tempFilePath)
        summary = CreateNewTextFile(summaryFilePath)
        tempFile = OpenTextFile_ReadOnly(tempFilePath)
        alert("UpdatingHeaders")
        var headerLine = tempFile.ReadLine()
        var numQuotes = headerLine.split('"').length - 1;
        if ((numQuotes % 2) == 1) // odd number of quotes means last escape is not unescaped
            headerLine += '"';
        var numNewHeaders = headers.length
        var fieldExtentions = ""
        for (var i = 0; i < numNewHeaders; i++) {
            headerLine += "," + headers[i] // add new headers
            fieldExtentions += "," // will add this to each existing data line to make sure numFields == numHeaders for existing data
        }
        summary.Write(headerLine)

        headerLine = SplitCsvLine(headerLine)
        headers = []
        for (var i = 0; i < headerLine.length; i++) {
            headerLine[i] = normalizeHeader(headerLine[i])
            headers[headerLine[i]] = i;
        }

        while (!tempFile.AtEndOfStream)
            summary.WriteLine(tempFile.ReadLine() + fieldExtentions)
        tempFile.Close()
        tempFile = null;
        DeleteFile(tempFilePath)
    } catch (e) {
        if (summary != null)
            summary.Close()
        throw e
    }finally {
        if (tempFile != null)
            tempFile.Close()
    }
    var result = {}
    result.summary = summary
    result.headers = headers
    return result
}
function normalizeHeader(header) {
    return header.replace(/^"+/, "").replace(/"+$/, "").toLowerCase();
}

function AddSummaryCreationMethodOption_CreateSummaryFile_MergeAll_EnforceHeaders() {
    var options = window.SummaryCreationMethodOptions
    if (options == null) {
        window.SummaryCreationMethodOptions = []
        options = window.SummaryCreationMethodOptions
    }
    var opt = {}
    opt.Name = "MergeAll_EnforceHeaders"
    opt.Description = "This will combine all the data from all the files into a single file, the summary file.\n" +
                       "This summary file, containing all the data from all the files, will be uploaded to the " +
                            "ftp site along with the files.\n" +
                       "If the header rows from some files do not match other files, additional columns will be " +
                            "added to support all columns from all files and the data will be sorted to be placed " +
                            "in the appropriate column with missing columns given empty fields.\n" +
                       "When a column is added, data from files which do not have that column will be given an empty " +
                            "value for that field.\n" +
                       "(for example: if file 1 has columns (this, that, andthensome) while file2 has columns (that, " +
                            "andthensome, andyetmore), the summary file will have columns (this, that, andthensome, " +
                            "andyetmore). the data from file 1 will have empty fields added at the end for the column " +
                            "(andyetmore) and the data from file2 will have empty fields added at the beginning for " +
                            "the column (this) )\n" +
                       "Note: only the header row of the summary file will contain all columns for all files, the data " +
                            "files themselves will not be affected."
    opt.Method = CreateSummaryFile_MergeAll_EnforceHeaders
    options.push(opt)
    options[opt.Name] = opt
}
AddSummaryCreationMethodOption_CreateSummaryFile_MergeAll_EnforceHeaders();