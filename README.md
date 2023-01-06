# TIFF Tools
This repository contains two free light-weight scripts - written in JScript - that can be used [to create and/or split multipage TIFF files](https://www.ilovefreesoftware.com/03/windows/image-photo/free-multipage-tiff-creator-to-create-multipage-tif-images.html). These scripts use [the Windows Image Acquisition (WIA) library v2.0](https://learn.microsoft.com/en-us/previous-versions/windows/desktop/wiaaut/-wiaaut-startpage), which is implemented by the file `wiaaut.dll`, located in the `C:\Windows\System32` folder. The second version of the WIA library that is required for the scripts is shipped with Windows Vista and later; so **it is not possible to run these scripts on Windows XP or older versions of Windows**.

## SplitTIFF.js
`SplitTIFF.js` is a light-weight script (bearly 5.85 KB in size) that takes a multi-page TIFF file as input and stores each of its pages into a separate single-page image file in one of these formats: **PNG, TIFF, JPEG, BMP, or GIF**  
It's such a free, easy-to-use TIFF Splitter utility!

To use it, open a Command Prompt window, navigate to the directory where `SplitTIFF.js` is stored, and then run the script  according to the following syntax:

    cscript splittiff.js SourceTIFFFile [DestPath] [/q:Quality]

Below is the description of the parameters:

- **`SourceTIFFFile`**: Required. The path of the TIFF file whose pages are to be splitted.
- **`DestPath`**: Optional. If specified, it must be set to the path of the resulting image files, each of which will contain one page of the original TIFF file. Regarding the syntax of this path, please note that the file name within this path must end with one of the supported extensions **(.png, .tif, .tiff, .jpg, .jpeg, .jpe, .jfif, .bmp, .dib, .gif)**, and must include the special placeholder `%d`, which is replaced during the operation with the number of the corresponding page in the original TIFF file.  
If not specified, a Browse dialog box will be automatically opened to let you select a folder in which the resulting image files are to be created. In this case, the resulting image files will be stored in PNG format; there is no way to choose another format via the Browse dialog box.
- **`Quality`**: Optional. If the `DestPath` parameter specifies that the resulting image files **be stored in JPEG format**, then this parameter specifies the quality of JPEG compression, as an integer in the range 1 to 100. This parameter can also affect the size of the resulting image files. Default value is 100.

### After running the command
After running the script with the aforementioned command, it generates some logs which are displayed in the Console window. The logs indicate the progress of the operation, as well as possible error messages. If the operation completes successfully, you will see the following message at the end of the logs:

    [notice] Done!

### Examples of usage
Say, on our Desktop, we have a TIFF file named `nature.tif` with 4 pages. Now we are going to split it and place each of its pages in the directory `F:\goodies` as a separate image file.

Hence the question, "In which format should these separate image files be?" Well, the format that works best in our script is PNG. **The PNG format** is suitable for the most common situations, though if we need to reduce the quality and size of the resulting image files, we must choose JPEG, which will be discussed later.

So in the Command Prompt window, enter the following command (notice the placeholder `%d` in the `DestPath` parameter):

    cscript splittiff.js C:\Users\Public\Desktop\nature.tif F:\goodies\%d.png

After the operation completes, head to the `F:\goodies` directory, where you will find 4 image files, named `1.png`, `2.png`, `3.png`, and `4.png`, respectively. So the file `1.png` contains the first page of our TIFF file, the file `2.png` contains the second, and so on. The preceding number in the file names is actually the number of the corresponding page in our TIFF file; and the placeholder `%d` that was written before `.png` in the command tells the script to insert the number of the page right before `.png` in the file name.

Anyway. Sometimes you don't have time to specify the `DestPath` parameter by typing `F:\goodies\%d.png` in the command. No problem! You can try omitting this part of the command and see what happens:

    cscript splittiff.js C:\Users\Public\Desktop\nature.tif

After running the command, a Browse dialog box will appear (like the image below). It will let you easiliy select your desired folder in which to store the resulting image files (in this example, `F:\goodies`). Then you can click OK to start the operation. Upon successful completion, you will have 4 image files in the Goodies folder.
