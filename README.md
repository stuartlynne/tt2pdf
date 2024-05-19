# tt2pdf
# Sun May 19 02:50:46 PM PDT 2024

This is a script to produce a Time Trial starting list for Time Trials being 
timed with CrossMgr and RaceDB.


## Usage

In RaceDB, go to the Time Trial and click on the "Start List" tab.  Click on the
"Export to Excel" button and save the file.

Using the Windows Explorer navigate to the folder where the file was saved.

Right click on the file and select "Open With".  Select "Choose another app".

Find the *tt2pdf.exe" file and select it. Do not check the "Always use this app"
as that would make it the default app for all Excel files.

The script will create a PDF file with the same name as the Excel file in the
same folder.

It will also send the file path to your browser which will open the PDF file.

Use the browser to print the PDF file.

## Installation

Clone this repository and copy the tt2pdf.exe file to where you want it.

## Usage

The Start list is a table of the riders with their start times with the following columns:

| --- | --- | --- | --- | --- | --- |
| Stopwatch | Bib | LastName | FirstName | Notes/Delay | Start Time |
| --- | --- | --- | --- | --- | --- |

- Stopwatch is the stopwatch time that the rider should start.
- Bib is the rider's bib number.
- LastName and FirstName are the rider's name.
- Notes/Delay is any notes or delay time for the rider.
- Start Time is the wall clock time the rider should start.

The Start Time is the official start time for the rider. This may be different from the 
Stopwatch time if the race was not started at the official start time for the race.

Note that the Start Time is the time the rider should start, not when they should arrive at the start line.

The Notes/Delay is to record any delay in the start of a rider due to a starter error. Generally this
would not be due to the rider being late to the start line.

As a general practice it is a good idea to leave open starts between categories. That allows for
a rider to be started out of order without having to change the start times of all the riders in the
category. Start them in an empty spot and pencil their information. 




