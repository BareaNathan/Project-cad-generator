# Project-cad-generator
A small Python script with Tkinter interface to automate manual manipulation of CSV and XLS in an electronics industry process.

This project comes with a simple solution for a manual process of manipulation of CSV and XLS files, the manipulation intends to create a new CSV file that contains
information to create a Pick and Place program for FUJI and MIRAE machines of some electronic boards in SMT process. This new CSV file needs to contain the
References of the components on the board, their respective positions in X, Y and Rotation from the Zero position (Fiducial mark), and the Partnumber of the SMD 
Component that the machine needs to put there.

As I said, handling these files was a manual job for the machine programmer, this job usually takes ~40 minutes depending on the size of the board. The machine
programmer needed to open the CSV file extracted from Altium CAD Design and manually read the Partnumber in the BOM (XLS) file for each reference and then write back to
this new CSV file this information, in the right order, so it can upload the file to the machine's programming interface.

The solution of these jobs started with a spreadsheet, where he could copy and paste the CSV file and the XLS file, then in another spreadsheet he had the output file. 
The spreadsheet reduces the average time of this process to ~15 min. But sometimes the spreadsheet was very slow, due to the complexity of the formulas. so i started
think of code that would be able to do this job in less time, and flexible for all the standard file formats used in this industry. Python was the
language used to do this, using some packages to manipulate the files, and Tkinter for the GUI, making everything easier. The code is not the prettiest code
you will see, but it worked in that particular process and reduced the average time to ~3 min, leaving the machine programmer free to continue programming
your board without complications and manual errors.
