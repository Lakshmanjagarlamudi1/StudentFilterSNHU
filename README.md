# Student Course Filtration Tool
## Overview
This Python application provides a graphical user interface (GUI) for filtering student course data from an Excel file based on specified subjects of interest. The filtered data is then saved to a new Excel file.

## Features
Input File Selection: Users can select the source Excel file containing student course data through a file dialog.

Output File Specification: Users can specify the location to save the filtered data as a new Excel file through a file dialog.

Subject Filtering: Users can specify subjects of interest by entering course IDs, and only students enrolled in those courses will be included                      in the filtered data.

Error Handling: The application provides error handling for cases such as invalid file formats or missing data.

## Dependencies
Python 3: The application is written in Python 3.

pandas: Used for data manipulation and analysis.

openpyxl: Required for working with Excel files.

tkinter: Utilized for creating the graphical user interface.
