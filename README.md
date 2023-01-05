# TIFF Tools
This repository contains two free light-weight scripts - written in JScript - that can be used [to create and/or split multipage TIFF files](https://www.ilovefreesoftware.com/03/windows/image-photo/free-multipage-tiff-creator-to-create-multipage-tif-images.html). These scripts use [the Windows Image Acquisition (WIA) library v2.0](https://learn.microsoft.com/en-us/previous-versions/windows/desktop/wiaaut/-wiaaut-startpage), which is implemented by the file `wiaaut.dll`, located in the `C:\Windows\System32` folder. The second version of the WIA library that is required for the scripts is shipped with Windows Vista and later; so **it is not possible to run these scripts on Windows XP or older versions of Windows**.

## SplitTIFF.js
`SplitTIFF.js` is a light-weight script (bearly 5.85 KB in size) that takes a multi-page TIFF file as input and stores each of its pages into a separate single-page image file in one of these formats: **PNG, TIFF, JPEG, BMP, or GIF**  
It's such a free, easy-to-use TIFF Splitter utility!

To use it, open a Command Prompt window, navigate to the directory where `SplitTIFF.js` is stored, and then run the script  according to the following syntax:

    cscript splittiff.js SourceTIFFFile [DestPath] [/q:Quality]

Below is the description of the parameters:

- **`SourceTIFFFile`**: Required. The path of the TIFF file whose pages are to be splitted.
- **`DestPath`**: Optional. If specified, it must be set to the path of the resulting image files, each of which will contain one page of the original TIFF file. Regarding the syntax of this path, please note that the file name within this path must end with one of the supported extensions **(.png, .tif, .tiff, .jpg, .jpeg, .jpe, .jfif, .bmp, .dib, .gif)**, and must include the special placeholder `%d`, which is replaced during the conversion operation with the number of the corresponding page in the original TIFF file.
