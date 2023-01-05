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

try {
    var tiff = WSH.CreateObject("WIA.ImageFile");
}
catch (err) {
    WSH.StdErr.WriteLine("[error] Unable to create an instance of the ImageFile object. Most likely you are running an old version of Microsoft Windows, which does not have Windows Image Acquisition (WIA) library v2.0 installed.");
    WSH.StdErr.WriteLine("  Technical details: " + err.message);
    WSH.StdErr.WriteLine("  Error code: " + err.number.toHRESULT());
    WSH.Quit(100);
}

var wiaFormatBMP = "{B96B3CAB-0728-11D3-9D7B-0000F81EF32E}";
var wiaFormatPNG = "{B96B3CAF-0728-11D3-9D7B-0000F81EF32E}";
var wiaFormatGIF = "{B96B3CB0-0728-11D3-9D7B-0000F81EF32E}";
var wiaFormatJPEG = "{B96B3CAE-0728-11D3-9D7B-0000F81EF32E}";
var wiaFormatTIFF = "{B96B3CB1-0728-11D3-9D7B-0000F81EF32E}";

WSH.StdOut.WriteLine("Welcome to SplitTIFF, the free light-weight TIFF splitter!");

var e = new Enumerator(WSH.Arguments.Unnamed);
if (e.atEnd()) {
    WSH.StdErr.WriteLine("[error] Not enough arguments. Missing source TIFF file name.");
    WSH.Quit(101);
}

var tiffName = e.item();
try {
    tiff.LoadFile(tiffName);
    if (tiff.FormatID != wiaFormatTIFF)
        throw new Error("The specified file is in an invalid format.");
}
catch (err) {
    WSH.StdErr.WriteLine("[error] Unable to load the TIFF file '" + tiffName + "'.");
    WSH.StdErr.WriteLine("  Technical details: " + err.message);
    WSH.StdErr.WriteLine("  Error code: " + ((err.number) ? err.number.toHRESULT() : "N/A"));
    WSH.Quit(102);
}

var fc = tiff.FrameCount;
WSH.StdOut.WriteLine("[notice] The source TIFF file was loaded with " + fc + " page(s).");

e.moveNext();
if (e.atEnd()) {
    var shell = WSH.CreateObject("Shell.Application");
    
    var BIF_EDITBOX = 0x00000010;
    var BIF_DONTGOBELOWDOMAIN = 0x00000002;
    var BIF_RETURNONLYFSDIRS = 0x00000001;
    
    WSH.StdOut.WriteLine("[notice] Opening a Browse dialog box to seek the destination path, which was not specified in the command line...");
    var folder = shell.BrowseForFolder(0, "Please select the folder where the pages of the TIFF file are to be stored as separate PNG image files.", BIF_EDITBOX | BIF_DONTGOBELOWDOMAIN | BIF_RETURNONLYFSDIRS);
    if (!folder)
        WSH.Quit(103);
    
    var destPath = fso.BuildPath(folder.Self.Path, "%d.png");
    WSH.StdOut.WriteLine("[notice] Selected destination path: '" + destPath + "'");
}
else
    var destPath = e.item();

if (destPath.indexOf("%d") == -1) {
    WSH.StdErr.WriteLine("[error] Missing the placeholder '%d' in the destination path.");
    WSH.Quit(104);
}

var ip = WSH.CreateObject("WIA.ImageProcess");
ip.Filters.Add(ip.FilterInfos("Convert").FilterID);

var namedArgs = WSH.Arguments.Named;
if (namedArgs.Exists("Q")) {
    var quality = Number(namedArgs("Q"));
    if (isNaN(quality))
        WSH.StdErr.WriteLine("[warn] The output image quality specified in the /q parameter is not a valid number. Proceeding with the default image quality (100)...");
    else {
        try {
            ip.Filters(1).Properties("Quality") = quality;
        }
        catch (err) {
            WSH.StdErr.WriteLine("[warn] Unable to set the quality of the output series of images to " + quality + ". Proceeding with the default image quality (100)...");
            WSH.StdErr.WriteLine("  Technical details: " + err.message);
            WSH.StdErr.WriteLine("  Error code: " + err.number.toHRESULT());
        }
    }
}

var outputFormat = fso.GetExtensionName(destPath).toLowerCase();

for (var i = 1;i <= fc;i++) {
    tiff.ActiveFrame = i;
    
    var frame;
    try {
        switch (outputFormat) {
        case "bmp" :
        case "dib" :
            ip.Filters(1).Properties("FormatID") = wiaFormatBMP;
            frame = ip.Apply(tiff);
            break;
        case "png" :
            ip.Filters(1).Properties("FormatID") = wiaFormatPNG;
            frame = ip.Apply(tiff);
            break;
        case "gif" :
            ip.Filters(1).Properties("FormatID") = wiaFormatGIF;
            frame = ip.Apply(tiff);
            break;
        case "jpg" :
        case "jpeg" :
        case "jpe" :
        case "jfif" :
            ip.Filters(1).Properties("FormatID") = wiaFormatJPEG;
            frame = ip.Apply(tiff);
            break;
        default : // TIFF
            ip.Filters(1).Properties("FormatID") = wiaFormatPNG;
            frame = ip.Apply(tiff);
            ip.Filters(1).Properties("FormatID") = wiaFormatTIFF;
            frame = ip.Apply(frame);
        }
    }
    catch (err) {
        WSH.StdErr.WriteLine("[warn] Unable to extract the page " + i + ".");
        WSH.StdErr.WriteLine("  Technical details: " + err.message);
        WSH.StdErr.WriteLine("  Error code: " + err.number.toHRESULT());
        continue;
    }
    
    try {
        frame.SaveFile(destPath.replace("%d", i));
    }
    catch (err) {
        WSH.StdErr.WriteLine("[warn] Unable to create an image file containing the page " + i + ".");
        WSH.StdErr.WriteLine("  Technical details: " + err.message);
        WSH.StdErr.WriteLine("  Error code: " + err.number.toHRESULT());
        continue;
    }
    
    WSH.StdOut.WriteLine("[notice] Extracted the page " + i + "/" + fc + ".");
}

WSH.StdOut.WriteLine("[notice] Done!");

function showError(message) {
    wshShell.Popup(message, 0, "SplitTIFF", 16);
}