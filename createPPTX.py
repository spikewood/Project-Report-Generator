'''
createPPTX creates a powerpoint from a given sample powerpoint.
The source will come from a project approval.
'''

# imports go here
from pptx import Presentation
# from pptx.util import Inches
import argparse
# import pandas as pd
from datetime import date

# Classes and globals
TITLE_SLIDE = {'layout': 0, 'subtitle': 1}
APPROVAL_SLIDE = {'layout': 2, 'subtitle': 1}


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


def create_ppt_approval_slide(prs, title_txt, data):
    '''Creates a project approval slide'''
    # choose the Title layout for the slide
    slide_layout = prs.slide_layouts[APPROVAL_SLIDE['layout']]
    slide = prs.slides.add_slide(slide_layout)
    # assign title text
    title = slide.shapes.title
    title.text = title_txt
    # assign subtitle text
    subtitle = slide.placeholders[APPROVAL_SLIDE['subtitle']]
    subtitle.text = data['subtitle']


def create_pptx(input, output):
    '''Take the input powerpoint file and use it as the template
    for the output file.'''

    prs = Presentation(input)

    # Titles slide
    title_txt = "Project Approvals"
    subtitle_txt = "Created on {:%m-%d-%Y}".format(date.today())
    create_ppt_title_slide(prs, title_txt, subtitle_txt)

    # Project Approvals Slides
    data = {'subtitle': "Created on {:%m-%d-%Y}".format(date.today())}
    create_ppt_approval_slide(prs, "Project Approvals", data)

    prs.save(output)


if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument("pptxSource", type=str,
                        help="source pptx")
    parser.add_argument("pptxDestination", type=str,
                        help="destination pptx")
    args = parser.parse_args()
    create_pptx(args.pptxSource, args.pptxDestination)
