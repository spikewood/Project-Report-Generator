'''
createPPTX creates a powerpoint from a given sample powerpoint.
The source will come from a project approval.
'''

# imports go here
from pptx import Presentation
# from pptx.util import Inches
import argparse
import pandas as pd
from datetime import date

# Classes and globals

# Placeholder layout and indices for the Title Slide
TITLE_SLIDE = {'layout': 0, 'subtitle': 1}

# Placeholder indices for the approval slide
APP_SLIDE = {'layout': 2, 'Description': 1, 'Status': 14,
             'Sponsor Department': 24, 'Return Type': 23,
             'Return on Investment': 15, 'Return': 16,
             'Total Investment': 17, }

# Data Column Names for simplified mapping
COLMUN_NAMES = {
           'Summary': 'Title',
           'Custom field (Sponsor Department)': 'Sponsor Department',
           'Custom field (Project Classification)': 'Project Classification',
           'Custom field (Return Type)': 'Return Type',
           'Custom field (Return on Investment)': 'Return on Investment',
           'Custom field (Break Even Period)': 'Break Even Period',
           'Custom field (Return)': 'Return',
           'Custom field (Duration)': 'Duration',
           'Custom field (Total Investment)': 'Total Investment',
           'Custom field (Labor T-Shirt Size)': 'Labor T-Shirt Size',
           'Custom field (Labor Investment)': 'Labor Investment',
           'Custom field (Non-Labor Investment)': 'Non-Labor Investment'
           }


# functions go here
def create_ppt_title_slide(prs, title_txt, subtitle_txt):
    '''Creates a title slide in the presentation'''
    # choose the Title layout for the slide
    slide_layout = prs.slide_layouts[TITLE_SLIDE['layout']]
    slide = prs.slides.add_slide(slide_layout)
    # assign title text
    title = slide.shapes.title
    title.text = title_txt
    # assign subtitle text
    subtitle = slide.placeholders[TITLE_SLIDE['subtitle']]
    subtitle.text = subtitle_txt


def placeTextInSlide(slide, placeholderIndex, text):
    '''Inserts text into a slide.
    Handles exceptions with blanks.'''
    try:
        slide.placeholders[placeholderIndex].text = text
    except TypeError:
        slide.placeholders[placeholderIndex].text = '0'


def create_ppt_approval_slide(prs, title_txt, data):
    '''Creates a project approval slide in the presentation'''

    # choose the Title layout for the slide
    slide_layout = prs.slide_layouts[APP_SLIDE['layout']]
    slide = prs.slides.add_slide(slide_layout)

    # assign title text
    slide.shapes.title.text = data["Title"]

    # assign data to placeholder text from the sample ppt
    placeTextInSlide(slide, APP_SLIDE['Description'], data['Description'])
    placeTextInSlide(slide, APP_SLIDE['Status'], data['Status'])
    placeTextInSlide(slide, APP_SLIDE['Sponsor Department'],
                     data['Sponsor Department'])
    placeTextInSlide(slide, APP_SLIDE['Return Type'], data['Return Type'])
    placeTextInSlide(slide, APP_SLIDE['Return on Investment'],
                     data['Return on Investment'])
    placeTextInSlide(slide, APP_SLIDE['Return'], data['Return'])
    placeTextInSlide(slide, APP_SLIDE['Total Investment'],
                     data['Total Investment'])


def getDataFrame(data_file):
    data_frame = pd.read_csv(data_file)
    data_frame.rename(columns=COLMUN_NAMES, inplace=True)
    return data_frame


def create_pptx(data_file, input_file, output_file):
    '''Take the input powerpoint file and use it as the template
    for the output file.'''

    prs = Presentation(input_file)

    # Titles slide
    title_txt = "Project Approvals"
    subtitle_txt = "Created on {:%m-%d-%Y}".format(date.today())
    create_ppt_title_slide(prs, title_txt, subtitle_txt)

    # get the data from the data file
    df = getDataFrame(data_file)

    # Project Approvals Slides
    create_ppt_approval_slide(prs, "Project Approvals", df.iloc[1])

    prs.save(output_file)


if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument("dataSource", type=str,
                        help="data source")
    parser.add_argument("pptxSource", type=str,
                        help="source pptx")
    parser.add_argument("pptxDestination", type=str,
                        help="destination pptx")
    args = parser.parse_args()
    create_pptx(args.dataSource, args.pptxSource, args.pptxDestination)
