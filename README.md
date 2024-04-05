# generate_ppt_from_word
Script for creating a PowerPoint presentation from a Word document.

## Overview:
This script facilitates the creation of a PowerPoint presentation from a Word document (.docx).  It prompts the user to select a Word file using a file dialog, to insert a title and subtitle for the portrait and then generates a Powerpoint file presentation with the Word document's content.

## Requirements:
Python 3
tkinter library
docx library
pptx library

## Files
generate_ppt_from_word.py

## Usage
1. Run the script.
2. A file dialog will prompt you to select a Word file.
3. After selecting the Word file, the script will ask you to insert the title and the subtitle of the presentation's portrait.
4. The resulting Powerpoint file will be saved in the same directory and have the same name as the Word file.

See "ppt_from_word_1.JPG" and "ppt_from_word_2".

## Important Note
Ensure that the selected file is a valid MS Word file in .docx format.
The script will use the first paragraph as Title and the second paragraph as Body. 
The third paragraph must be left empty, as it serves as a separator between this slide and the following one.
Ensure the Word file comply with the structure:
  Title 1
  Body 1

  Title 2
  Body 2

See "ppt_from_word_3"

## License
This project is governed by the GNU Affero General Public License v3.0. For comprehensive details, kindly refer to the LICENSE file included with this project.
