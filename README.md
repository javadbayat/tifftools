# TIFF Tools
This repository contains two free light-weight scripts - written in [JScript](http://msdn.microsoft.com/library/hbxc2t98.aspx) - that can be used [to create and/or split multipage TIFF files](https://www.ilovefreesoftware.com/03/windows/image-photo/free-multipage-tiff-creator-to-create-multipage-tif-images.html). These scripts use [the Windows Image Acquisition (WIA) library v2.0](https://learn.microsoft.com/en-us/previous-versions/windows/desktop/wiaaut/-wiaaut-startpage), which is implemented by the file `wiaaut.dll`, located in the `C:\Windows\System32` folder. The second version of the WIA library that is required for the scripts is shipped with Windows Vista and later; so **it is not possible to run these scripts on Windows XP or older versions of Windows**.

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

![BrowseForFolder](https://user-images.githubusercontent.com/31417320/211037195-93791fe7-be1e-465a-8b8f-9033090c0cb1.jpg)

By default, the resulting image files are created in full quality. If you want to reduce the quality of the resulting image files so that their size also decreases, then set their format to JPEG instead of PNG, and append `/q:` to the command line, followed by an integer between 1 and 100 which indicates your desired image quality. For example, the following command produces some JPEG image files with their quality set to 40.

    cscript splittiff.js C:\Users\Public\Desktop\nature.tif F:\goodies\%d.jpg /q:40

### My final wish
> Wish one day my program will be introduced as an answer to [this SuperUser post](https://superuser.com/questions/44600/how-to-split-a-multipage-tiff-file-on-windows).

## MakeTIFF.js
`MakeTIFF.js` is a light-weight script (bearly 7.61 KB in size) that takes several regular single-page image files as input, and merges them into one multi-page TIFF file. For the input image files, the supported formats are **PNG, TIFF, JPEG, BMP, and GIF**.  
It's such a free, easy-to-use TIFF Creation utility! It is actually the inverse of `SplitTIFF.js` script.

You can use this script in two modes: **Normal Mode** or **Browse Mode**. To use the script, you must first open a Command Prompt window and navigate to the directory where `MakeTIFF.js` is stored. Then read the below explainations for your desired mode, and follow the instructions.

### Using in Normal Mode
In the Command Prompt, execute `cscript maketiff.js`, followed by the space-delimited list of the input image files, followed by `/o`, followed by the path of the output TIFF file that is to be created. For example, the following command creates a multipage TIFF file containing three images from the `F:\exam` directory and three images from the `F:\river` directory. It then stores the resulting TIFF file into the path `F:\collection.tiff`.

    cscript maketiff.js F:\exam\answersheet_alice.bmp F:\exam\answersheet_bob.bmp F:\exam\answersheet_emma.bmp F:\river\photo1.jpg F:\river\photo2.jpg F:\river\photo3.jpg /o F:\collection.tiff

What is interesting is that each item in the list of input image files may include wildcards (`*` or `?`). So with the use of wildcards, the preceding command can be simplified as follows:

    cscript maketiff.js F:\exam\answersheet_*.bmp F:\river\photo?.jpg /o F:\collection.tiff

By default, the resulting TIFF file is compressed using the **LZW** algorithm. If you wish to change the compression scheme of the output TIFF file, then you can append `/c:` to the command, followed by one of the supported compression schemes (**`CCITT3`, `CCITT4`, `RLE`, or `LZW`**), or append **`/c:0` or `/c:Uncompressed`**, so the output TIFF file will not be compressed at all. For example, the following command sets the compression scheme of the resulting TIFF file to `RLE`:

    cscript maketiff.js F:\exam\answersheet_*.bmp F:\river\photo?.jpg /o F:\collection.tiff /c:RLE

And the following command disables the compression of the resulting TIFF file:

    cscript maketiff.js F:\exam\answersheet_*.bmp F:\river\photo?.jpg /o F:\collection.tiff /c:0

**Note:** The compression schemes that can be specified with the `/c` parameter are case-sensitive.

### Using in Browse Mode
*If you consider the Normal Mode too difficult to use, then this mode provides an easier way to create a multipage TIFF file.*

Open one or more File Explorer windows, and select the desired input image files there. Then while the File Explorer windows are open with the input image files being in the highlighted state, switch to the Command Prompt window and execute the following command:

    cscript maketiff.js /b /o OutputTIFFFile

where `OutputTIFFFile` is the path of the output multipage TIFF file that is to be created. For example, the following command stores the output TIFF file into the path `F:\collection.tiff`:

    cscript maketiff.js /b /o F:\collection.tiff

After executing the command, the script will automatically examine all those File Explorer windows and detect the selected input image files. Then it will merge the input image files into one multipage TIFF file, which will be stored in the path `F:\collection.tiff`.
