# Loop Internet PoleTool Overview
The Loop Internet PoleTool is an all-in-one program which provides all the tools necessary to parse pole data and prepare/export it in different forms to help make the Design/GIS team's job easier!

## Features *(Latest v1.6)*
- ### Prepare for Vetro
  Refactors and formats pole data to a new Excel sheet such that when it is imported to Vetro there is no need to match up attributes manually. This action also helps with importing data into QGIS.
- ### Generate Make Ready Notes
  Generates a Make Ready Notes Excel sheet which is typically submitted alongside a strand map in Monday.com.
- ### Generate Verizon Application
  Generates a Verizon Pole Application Excel sheet which is formatted such that it can be submitted to Verizon right away. Turning ON the Use API switch will increase the time needed to complete the operation but this will split the poles by municipality and export them into separate Excel sheets.
- ### Generate Frontier Application (Work in progress)
  Currently under development. Stay tuned for updates!

# How To Use
Using the PoleTool is simple and straightforward. Let's dive in!

![app.png](docs/app.png)

On the right side of the application is your action checklist. From here you can select one or more file operations you'd like to execute on your input file. On the right side of the application is the file management zone and start button. From here you will be able to input a file, select the ouput for your newly generated file(s), and name the output files.

## Steps:
1. Drop your node attributes Excel sheet you got from Katapult into the file drag and drop area at the top of the right side.
2. Click the Select Output button to choose the output location for your new files.
3. Enter in an output file name if desired. This will default to the same filename as your inputted filename.
4. Select which operation(s) you would like to perform on the file.
5. Click Process File and wait for the progress bar to say complete.
6. Done! Check that your files were outputted properly.

### Note: Input file should be node attributes xlsx. Not with ID's or any other file format from katapult. This may cause data to be read improperly.
