function CreateSummaryFile_MergeAll_EnforceHeaders(filePath, folderPath) {

    throw new Error("Not Tested")


    var files = GetFiles(txt.value);
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
        file = OpenTextFile_ReadOnly(files.pop())
        var headerLine = SplitCsvLine(file.ReadLine())
        var headers = []
        for (var i = 0; i < headerLine.length; i++) {
            headerLine[i] = normalizeHeader(headerLine[i])
            headers[headerLine[i]] = i;
        }
        while( !file.AtEndOfStream )
            summary.WriteLine(file.ReadLine())
        file.Close()
        file = null
        while (files.length > 0) {
            var lineCount = 0
            var currFilePath = files.pop()
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
                }
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
    } finally {
        if (summary != null)
            summary.Close()
        if (file != null)
            file.Close()
    }
}
function UpdateHeaders(summary, headers, summaryFilePath) {    
    var tempFile
    try{
        summary.Close()
        summary = null
        var tempFilePath = summaryFilePath + ".temp"
        MoveFile(summaryFilePath, tempFilePath)
        summary = CreateNewTextFile(summaryFilePath)
        tempFile = OpenTextFile_ReadOnly(tempFilePath)

        var headerLine = tempFile.ReadLine
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
    } finally {
        if (summary != null)
            summary.Close()
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