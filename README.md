# openlawnz-pdf-extraction

## Requirements

* Microsoft Office 365 (other older versions may work)
* Visual Studio 2019
* [Sample case files in a folder](https://openlawnz-my.sharepoint.com/:f:/g/personal/andrew_openlaw_nz/EvknIZ3w4YdGvHVeRg5LVFsBDzvZdjvmm3TKbh9OMSLJGw?e=eLlbIb) (which you will point to - download this folder)

## Running

1. Open the Solution File
2. Run the `WordToText` project
3. Point to the sample cases folder

## Debugging

1. Put sample files in `__openlawnz_from_pdf` that you want to process (it needs that string in the path)
2. Open the Solution File
3. Run the `CaseDataExtractor` projet
4. Choose a Word file in the folder

## Output

Nothing will show, but 3 text files will appear next to the source file.

* `filename.txt`
* `filename.footnotecontexts.txt` 
* `filename.footnotes.txt`

A log file per run is also generated prefixed `_log`

## Things to note

* Sometimes there are leftover Word processes and files. The processor will clean it up before it runs each time, but if you use MS Word, it may intefere, so open up Task Manager and kill any old Word processes.
