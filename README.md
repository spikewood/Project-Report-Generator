# create_ppt.py
createPPTX is a set of python scripts designed to create a powerpoint slide deck.

This project will build the following into it:
1. A method to determine the layout of the source createPPTX
1. A method to read in a csv file into a slide
1. A way to publish the slides

## Dependencies
You will need to have both *pptx* and *pandas* modules to run these scripts.

Installing python-pptx through the terminal:
`pip install python-pptx`

Installing pandas through the terminal:
`pip install pandas`

## Analyzing PPT files
You can analyze the placeholder structure of a pptx by utilizing the analyze_ppt.py script. Example:

`python3 analyze_ppt.py ppt_sample.pptx ppt_analysis.pptx`

The result will show how the master slides layout the placeholders including the indexes of the placeholders.

_**Thanks to Chris Moffit from chris1610/pbpython for this script and his atricle http://pbpython.com/creating-powerpoint.html_

## Creating a ppt
You can create a pptx deck utilizing the create_ppt.py script. Example:

`python3 create_ppt.py sample_data.csv ppt_sample.pptx ppt_output.pptx`


