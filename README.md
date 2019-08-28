# create_ppt.py
createPPTX is a set of python scripts designed to create a powerpoint slide deck.

This project will build the following into it:
1. A method to determine the layout of the source createPPTX
1. A method to read in a csv file into a slide
1. A way to publish the slides

## Analyzing PPT files
You can analyze the placeholder structure of a pptx by utilizing the analyze_ppt.py script. Example:

`python3 analyze_ppt.py ppt_sample.pptx ppt_analysis.pptx`

The result will show how the master slides layout the placeholders including the indexes of the placeholders.

## Creating a ppt
You can create a pptx deck utilizing the create_ppt.py script.
