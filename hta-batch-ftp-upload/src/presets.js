window.Presets_js = {};

function AddPreset_js(bAddToUI, Name, Directory, SummaryFilename, FtpHost, FtpDirectory, FtpUsername, FtpPassword) {
    var result = {};
    result.Name = Name
    result.Directory = Directory
    result.SummaryFilename = SummaryFilename
    result.FtpHost = FtpHost
    result.FtpDirectory = FtpDirectory
    result.FtpUsername = FtpUsername
    result.FtpPassword = FtpPassword
    window.Presets_js[result.Name] = result;
    if (bAddToUI)
        AddPresetOption(Name, Name);
}