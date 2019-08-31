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
TITLE_SLIDE_LAYOUT = 0
TITLE_SLIDE = {'subtitle': 1}

BLANK_SLIDE_LAYOUT = 7

# Placeholder indices for the first approval slide
APP_SLIDE1_LAYOUT = 2
APP_SLIDE1 = {'Key': 27,
              'Department': 23,
              'Description': 1,
              'Status': 24,
              'Return Type': 25,
              'ROI': 15,
              'Return': 16,
              'Total Investment': 17,
              'Sponsor': 19,
              'Project Owner': 18,
              'Line of Business': 21,
              'Return Description': 28,
              'Investment Description': 13,
              }

# Placeholder indices for the second approval slide
APP_SLIDE2_LAYOUT = 3
APP_SLIDE2 = {'Scope': 14,
              'Scope Exclusions': 15,
              'Systems': 17,
              'Business Area Impacts': 16,
              }

# Data Column Names for simplified mapping
COLMUN_NAMES = {'Issue key': 'Key',
                'Issue id': 'ID',
                'Created': 'Created',
                'Summary': 'Summary',
                'Custom field (Sponsor Department)': 'Department',
                'Custom field (Project Classification)':
                    'Classification',
                'Description': 'Description',
                'Status': 'Status',
                'Custom field (Return Type)': 'Return Type',
                'Custom field (Return on Investment)': 'ROI',
                'Custom field (Break Even Period)': 'Break Even Period',
                'Custom field (Return)': 'Return',
                'Custom field (Duration)': 'Duration',
                'Custom field (Total Investment)': 'Total Investment',
                'Custom field (Labor T-Shirt Size)': 'Labor T-Shirt Size',
                'Custom field (Labor Investment)': 'Labor Investment',
                'Custom field (Non-Labor Investment)': 'Non-Labor Investment',
                'Custom field (Executive Sponsor)': 'Sponsor',
                'Custom field (Sponsor)': 'Project Owner',
                'Custom field (Business Area Impacts)':
                    'Business Area Impacts',
                'Custom field (Labor Investment Description)':
                    'Investment Description',
                'Custom field (Line of Business)': 'Line of Business',
                'Custom field (Return Description)': 'Return Description',
                'Custom field (Scope)': 'Scope',
                'Custom field (Scope Exclusions)': 'Scope Exclusions',
                'Custom field (Systems)': 'Systems',
                }

STATUS_MAPPING = {"Project Proposals": ["New Request",
                                        "Initial Review",
                                        "More Details Required",
                                        "Executive Sponsor Review",
                                        "Estimate Costs",
                                        "ROI Validation",
                                        "Executive Committee Review"],
                  "Inflight Projects": ["Project Approved",
                                        "Project Scheduled",
                                        "Project Underway",
                                        "Project On Hold"],
                  "Complete Projects": ["Project Complete"],
                  "Rejected Projects": ["Project Rejected"],
                  }


def createSlide(prs, layout_index):
    '''returns a new slide based on the slide index'''
    # add the slide from based on the layout
    slide_layout = prs.slide_layouts[layout_index]
    return prs.slides.add_slide(slide_layout)


def populateSlideFromList(slide, title_txt, placeholder_list):
    '''Populates a slide in the presentation based on a title,
    and a list of placeholder index and text pairs.
    Empty text will skip insertion of that element.'''
    # insert the title into the slide
    if title_txt:
        slide.shapes.title.text = title_txt
    # place the values into the slide
    for placeholder_index, txt in placeholder_list:
        placeTextInSlide(slide, placeholder_index, txt)


def populateSlideFromSeries(slide, title_txt, placeholder_map, series):
    '''Populates a slide based on the map of placeholders to a dataframe series.
    Ensures the data_frame contains the necessary columns. Unknown columns are
    removed.'''
    # insert the title into the slide
    if title_txt:
        slide.shapes.title.text = title_txt
    # reduce the placeholder map down to the columns that exsist in the series
    reduced_placeholder_map = {column_name: placeholder_map[column_name]
                               for column_name in placeholder_map.keys()
                               if column_name in series.keys()}
    # map the data columns to the placeholders
    for column_name, placeholder_index in reduced_placeholder_map.items():
        placeTextInSlide(slide, placeholder_index, series[column_name])


def placeTextInSlide(slide, placeholderIndex, text):
    '''Inserts text into a slide.
    Handles exceptions with blanks.'''
    try:
        slide.placeholders[placeholderIndex].text = text
    except TypeError:
        slide.placeholders[placeholderIndex].text = 'Not Available'


def getDataFrame(data_file):
    data_frame = pd.read_csv(data_file)
    data_frame.rename(columns=COLMUN_NAMES, inplace=True)
    return data_frame


def create_pptx(data_file, input_file, output_file):
    '''Take the input powerpoint file and use it as the template
    for the output file.'''

    # get the data from the data file
    print('Creating data frame.')
    df = getDataFrame(data_file)

    for dept in df['Department'].unique():
        # build a presentation for each of the departments
        prs = Presentation(input_file)

        # Insert title slide with department and status_category
        print('Creating', dept, 'title slide.')
        for status_category in STATUS_MAPPING:
            statuses = STATUS_MAPPING[status_category]
            df_subset = df.loc[df['Department'] == dept]
            df_subset = df_subset.loc[df['Status'].isin(statuses)]
            if not df_subset.empty:
                title_txt = dept + ' ' + status_category
                subtitle_txt = (dept + ' ' + status_category +
                                '\nCreated on {:%m-%d-%Y}'.format(date.today())
                                )
                placeholder_list = [(TITLE_SLIDE['subtitle'], subtitle_txt)]
                populateSlideFromList(createSlide(prs, TITLE_SLIDE_LAYOUT),
                                      title_txt, placeholder_list)
                # Insert a blank slide
                createSlide(prs, BLANK_SLIDE_LAYOUT)

                # create the slides in status order
                for status in STATUS_MAPPING[status_category]:
                    # subset the dataframe to the department rows
                    print(dept, status)
                    df_subset = df.loc[(df['Department'] == dept) &
                                       (df['Status'] == status)]
                    df_subset = df_subset.sort_values(by=['Summary'])
                    # Proposed Project Slides
                    for index, series in df_subset.iterrows():
                        print('Creating approval slide ', index + 1, 'of',
                              len(df_subset.index), '.')
                        title_txt = series["Summary"]
                        populateSlideFromSeries(createSlide(prs,
                                                            APP_SLIDE1_LAYOUT),
                                                title_txt, APP_SLIDE1, series)
                        populateSlideFromSeries(createSlide(prs,
                                                            APP_SLIDE2_LAYOUT),
                                                title_txt, APP_SLIDE2, series)

        dept_output_file = output_file.replace('.pptx', '_' +
                                               str(dept).replace(' ', '_') +
                                               '.pptx')
        prs.save(dept_output_file)


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
