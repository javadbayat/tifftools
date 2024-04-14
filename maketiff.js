Number.prototype.toHRESULT = function () {
    var hr = (0x100000000 + this).toString(16).toUpperCase();
    if (hr.length > 8)
        throw new Error("Number too large for HRESULT.");
    while (hr.length != 8)
        hr = "0" + hr;
    return ("0x" + hr);
};

var fso = WSH.CreateObject("Scripting.FileSystemObject");
var wshShell = WSH.CreateObject("WScript.Shell");

if (fso.GetBaseName(WSH.FullName).toLowerCase() != "cscript") {
    showError("This program must be run using 'cscript.exe'.");
    WSH.Quit(1);
}

var outputFileName = "";
var baseFrame = null;
var shell = WSH.CreateObject("Shell.Application");

try {
    var ip = WSH.CreateObject("WIA.ImageProcess");
}
catch (err) {
    WSH.StdErr.WriteLine("[error] Unable to create an instance of the ImageProcess object. Most likely you are running an old version of Microsoft Windows, which lacks Windows Image Acquisition (WIA) library v2.0.");
    WSH.StdErr.WriteLine("  Technical details: " + err.message);
    WSH.StdErr.WriteLine("  Error code: " + err.number.toHRESULT());
    WSH.Quit(100);
}

WSH.StdOut.WriteLine("Welcome to MakeTIFF, the free light-weight TIFF maker!");

var e = new Enumerator(WSH.Arguments);
for (;!e.atEnd();e.moveNext()) {
    var arg = e.item();
    if ((arg == "/o") || (arg == "/O")) {
        e.moveNext();
        if (!e.atEnd())
            outputFileName = e.item();
        
        break;
    }
    
    if ((arg == "/b") || (arg == "/B")) {
        browseForImageFiles();
        continue;
    }
    
    if (arg.charAt(0) == "/")
        continue;
    
    var fileName = fso.GetFileName(arg);
    var containsWC = (fileName.indexOf("*") != -1);
    var containsQM = (fileName.indexOf("?") != -1);
    if (containsWC && containsQM) {
        WSH.StdErr.WriteLine("[warn] Wildcards (*) and question marks (?) must not be both used within a file name pattern.");
        WSH.StdErr.WriteLine("  The given file name pattern was '" + fileName + "'.");
    }
    else if (containsWC) {
        var folderName = fso.GetAbsolutePathName(fso.GetParentFolderName(arg));
        if (fso.FolderExists(folderName)) {
            var folder = fso.GetFolder(folderName);
            var rule = fileName.toLowerCase().split("*");
            
            var ef = new Enumerator(folder.Files);
            for (;!ef.atEnd();ef.moveNext()) {
                var file = ef.item();
                if (wildcardMatches(file.Name.toLowerCase(), rule))
                    loadFrame(file.Path);
            }
        }
        else
            WSH.StdErr.WriteLine("[warn] Unable to find the folder '" + folderName + "'.");
    }
    else if (containsQM) {
        var folderName = fso.GetAbsolutePathName(fso.GetParentFolderName(arg));
        if (fso.FolderExists(folderName)) {
            var folder = fso.GetFolder(folderName);
            var rule = fileName.toLowerCase();
            
            var ef = new Enumerator(folder.Files);
            for (;!ef.atEnd();ef.moveNext()) {
                var file = ef.item();
                if (qmMatches(file.Name.toLowerCase(), rule))
                    loadFrame(file.Path);
            }
        }
        else
            WSH.StdErr.WriteLine("[warn] Unable to find the folder '" + folderName + "'.");
    }
    else
        loadFrame(arg);
}

if (!outputFileName) {
    WSH.StdErr.WriteLine("[error] Not enough arguments. Missing output image file name.");
    WSH.Quit(101);
}

if (!baseFrame) {
    WSH.StdErr.WriteLine("[error] No image files were loaded as the frames of the TIFF file.");
    WSH.Quit(102);
}

WSH.StdOut.WriteLine("[notice] Converting to TIFF...");

var wiaFormatTIFF = "{B96B3CB1-0728-11D3-9D7B-0000F81EF32E}";

ip.Filters.Add(ip.FilterInfos("Convert").FilterID);
ip.Filters(ip.Filters.Count).Properties("FormatID") = wiaFormatTIFF;

var namedArgs = WSH.Arguments.Named;
if (namedArgs.Exists("C")) {
    var compression = namedArgs("C");
    if (compression == "0")
        compression = "Uncompressed";
    
    try {
        ip.Filters(ip.Filters.Count).Properties("Compression") = compression;
    }
    catch (err) {
        WSH.StdErr.WriteLine("[warn] Unable to set the compression scheme to '" + compression + "'. Creating the resulting TIFF file with the default compression setting (LZW) anyway...");
        WSH.StdErr.WriteLine("  Technical details: " + err.message);
        WSH.StdErr.WriteLine("  Error code: " + err.number.toHRESULT());
    }
}

try {
    var tiff = ip.Apply(baseFrame);
}
catch (err) {
    WSH.StdErr.WriteLine("[error] Unable to convert the image files to TIFF.");
    WSH.StdErr.WriteLine("  Technical details: " + err.message);
    WSH.StdErr.WriteLine("  Error code: " + err.number.toHRESULT());
    WSH.Quit(103);
}

WSH.StdOut.WriteLine("[notice] Creating image file '" + outputFileName + "'...");

try {
    tiff.SaveFile(outputFileName);
}
catch (err) {
    WSH.StdErr.WriteLine("[error] Unable to create the output TIFF file.");
    WSH.StdErr.WriteLine("  Technical details: " + err.message);
    WSH.StdErr.WriteLine("  Error code: " + err.number.toHRESULT());
    WSH.Quit(104);
}

WSH.StdOut.WriteLine("[notice] Done!");

function showError(message) {
    wshShell.Popup(message, 0, "MakeTIFF", 16);
}

function loadFrame(fileName) {
    var frame = WSH.CreateObject("WIA.ImageFile");
    
    try {
        frame.LoadFile(fileName);
    }
    catch (err) {
        WSH.StdErr.WriteLine("[warn] Unable to load the image file '" + fileName + "'.");
        WSH.StdErr.WriteLine("  Technical details: " + err.message);
        WSH.StdErr.WriteLine("  Error code: " + err.number.toHRESULT());
        return false;
    }
    
    if (baseFrame) {
        ip.Filters.Add(ip.FilterInfos("Frame").FilterID);
        ip.Filters(ip.Filters.Count).Properties("ImageFile") = frame;
    }
    else
        baseFrame = frame;
    
    WSH.StdOut.WriteLine("[notice] Loaded image file '" + fileName + "'.");
    return true;
}

function wildcardMatches(str, rule) {
    var currentLoc = 0;
    for (var i = 0; i < rule.length; i++) {
        var loc = str.indexOf(rule[i], currentLoc);
        if (loc < 0)
            return false;
        if ((!i) && (rule[0] != ""))
            if (loc)
                return false;
        if ((i == rule.length - 1) && (rule[i] != ""))
            if (loc + rule[i].length != str.length)
                return false;
        currentLoc = loc + rule[i].length;
    }
    return true;
}

function qmMatches(str, rule, questionMark) {
    if (!questionMark)
        questionMark = "?";
    
    if (str.length != rule.length)
        return false;
    
    for (var i = 0; i < rule.length; i++) {
        if (rule.charAt(i) != questionMark) {
            if (rule.charAt(i) != str.charAt(i))
                return false;
        }
    }
    
    return true;
}

function browseForImageFiles() {
    WSH.StdOut.WriteLine("[notice] Looking for selected items in the current File Explorer windows...");
    
    var wins = shell.Windows();
    for (var i = 0;i < wins.Count;i++) {
        var win = wins.Item(i);
        if (win) {
            if (win.Name == "Windows Explorer") {
                var selectedFiles = win.Document.SelectedItems();
                for (var j = 0;j < selectedFiles.Count;j++) {
                    var fileName = selectedFiles.Item(j).Path;
                    if (fso.FileExists(fileName))
                        loadFrame(fileName);
                }
            }
        }
    }
    
    WSH.StdOut.WriteLine("[notice] Finished scanning File Explorer windows.");
}
