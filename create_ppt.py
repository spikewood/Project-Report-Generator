'''
createPPTX creates a powerpoint from a given sample powerpoint.
The source will come from a project approval.
'''

# imports go here
from pptx import Presentation
# from pptx.util import Inches
import argparse
import pandas as pd
import numpy as np
import os
from datetime import date

# Classes and globals

# Placeholder layout and indices for the Title Slide
TITLE_SLIDE_LAYOUT = 0
TITLE_SLIDE = {'subtitle': 1}

# Placeholder layout for blank slides
BLANK_SLIDE_LAYOUT = 7

# Placeholder indices for the first approval slide
PAP_SLIDE1_LAYOUT = 2
PAP_SLIDE1_COLUMN_MAPPING = {'Key': 27,
                             'Sponsor Department': 23,
                             'Description': 1,
                             'Status': 24,
                             'Return Type': 25,
                             'ROI': 15,
                             'Return': 16,
                             'Total Investment': 17,
                             'Sponsor': 18,
                             'Line of Business': 21,
                             'Return Description': 28,
                             'Investment Description': 13,
                             'Project Manager': 20,
                             'States': 22,
                             }

# Placeholder indices for the second approval slide
PAP_SLIDE2_LAYOUT = 3
PAP_SLIDE2_COLUMN_MAPPING = {'Key': 27,
                             'Scope': 14,
                             'Scope Exclusions': 15,
                             'Systems': 17,
                             'Business Area Impacts': 16,
                             }

# Data Column Names for simplified mapping
COLUMNS = {'Issue key': 'Key',
           'Issue id': 'ID',
           'Created': 'Created',
           'Summary': 'Summary',
           'Custom field (Sponsor Department)': 'Sponsor Department',
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
           'Custom field (Sponsor)': 'Sponsor',
           'Custom field (Business Area Impacts)': 'Business Area Impacts',
           'Custom field (Labor Investment Description)':
               'Investment Description',
           'Custom field (Line of Business)': 'Line of Business',
           'Custom field (Return Description)': 'Return Description',
           'Custom field (Scope)': 'Scope',
           'Custom field (Scope Exclusions)': 'Scope Exclusions',
           'Custom field (Systems)': 'Systems',
           'Custom field (Project Manager)': 'Project Manager',
           }

STATUS_MAPPING = {"Project Requests": ["New Request",
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

COLUMNS_WITH_NAMES = ['Project Manager', 'Sponsor']


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
        text_lines = str(text).splitlines()
        line_count = 0
        for line in text_lines:
            if line_count == 0:
                if line == 'nan':
                    slide.placeholders[placeholderIndex].text_frame.text = ''
                else:
                    slide.placeholders[placeholderIndex].text_frame.text = line
            elif line != '':
                paragraph = slide.placeholders[placeholderIndex] \
                            .text_frame.add_paragraph()
                paragraph.text = line
            line_count += 1
    except TypeError:
        slide.placeholders[placeholderIndex].text = 'Not Available'
    except KeyError:
        return


def convert_usernames_to_fullnames(data_frame):
    ''' compares a set of columnns against known usernames and converts them
        to full names '''
    # get the dictionary from the name_mapping.csv
    name_df = pd.read_csv("name_mapping.csv")
    name_dict = {name_df['Username'][ind]: name_df['Full Name'][ind]
                 for ind in name_df.index}
    for column in COLUMNS_WITH_NAMES:
        data_frame[column] = data_frame[column].map(name_dict)
    return data_frame


def reduce_data_frame_columns(data_frame):
    ''' reduces the columns to just those needed for the pap'''
    # get the needed column names for the slide decks
    needed_columns = (PAP_SLIDE1_COLUMN_MAPPING.keys() +
                      PAP_SLIDE2_COLUMN_MAPPING.keys())
    # filter the data_frame to just those needed columnns
    data_frame = data_frame.filter(needed_columns.unique())
    return data_frame


def clean_data_frame(data_frame):
    ''' Cleans the data in the data_frame. '''
    # Update column names in the data_frame
    data_frame.rename(columns=COLUMNS, inplace=True)
    # reduce the columns to those needed
    data_frame = reduce_data_frame_columns(data_frame)
    # change usernames to Full Names
    data_frame = convert_usernames_to_fullnames(data_frame)
    return data_frame


def getDataFrame(data_file):
    data_frame = pd.read_csv(data_file)
    data_frame = clean_data_frame(data_frame)
    return data_frame


def create_pap_pptx(df, prs, item, output_filename):
    # Insert title slide with department and status_category such as
    #    project request, inflight or complete.
    for status_category in STATUS_MAPPING:
        statuses = STATUS_MAPPING[status_category]
        status_dataframe_subset = df.loc[df['Status'].isin(statuses)]
        if not status_dataframe_subset.empty:
            title_txt = item + ' ' + status_category
            subtitle_txt = ('\nCreated on {:%m-%d-%Y}'.format(date.today())
                            )
            placeholder_list = [(TITLE_SLIDE['subtitle'], subtitle_txt)]
            populateSlideFromList(createSlide(prs, TITLE_SLIDE_LAYOUT),
                                  title_txt, placeholder_list)
            # Insert a blank slide
            createSlide(prs, BLANK_SLIDE_LAYOUT)

            # create the slides in status order
            for status in STATUS_MAPPING[status_category]:
                # subset the dataframe to the department rows
                df_subset = df.loc[(df['Status'] == status)]
                df_subset = df_subset.sort_values(by=['Summary'])
                i = 0
                for index, series in df_subset.iterrows():
                    if (len(df_subset.index) > 0) and (i == 0):
                        print(item, '-', status)
                        # Proposed Project Slides
                    i += 1
                    print('Creating slide', i, 'of',
                          len(df_subset.index), '.')
                    title_txt = series["Summary"]

                    first_pap_slide = createSlide(prs, PAP_SLIDE1_LAYOUT)
                    populateSlideFromSeries(first_pap_slide, title_txt,
                                            PAP_SLIDE1_COLUMN_MAPPING, series)

                    second_pap_slide = createSlide(prs, PAP_SLIDE2_LAYOUT)
                    populateSlideFromSeries(second_pap_slide, title_txt,
                                            PAP_SLIDE2_COLUMN_MAPPING, series)

        prs.save(output_filename)


def create_pap_pptxs(data_frame, ppt_template, file_prefix):
    '''Create a project approval ppt from a dataframe and ppt template.
        A pap ppt deck will be created for each dept with the output prefix.'''

    # make a folder for the ppts
    ppt_folder_path = file_prefix.replace('.pptx', '')
    if not os.path.exists(ppt_folder_path):
        os.makedirs(ppt_folder_path)

    # create a base path and filename for the ppts
    ppt_base_file_path = os.path.join(
                        ppt_folder_path,
                        os.path.basename(file_prefix))

    # slice determines the breakdown of project requests into ppts.
    slice = 'Sponsor'

    # any slice that is blank or nan should be marked unknown
    data_frame[slice] = data_frame[slice].replace(np.nan, 'Unknown',
                                                  regex=True)

    # create a ppt deck for each department
    for item in data_frame[slice].unique():
        # build a ppt deck for each of the departments
        ppt_deck = Presentation(ppt_template)

        # create the output filename using the file_prefix and department
        output_ppt_filename = ppt_base_file_path.replace(
                                '.pptx', '_' + str(item).replace(' ', '_') +
                                '.pptx')
        # get the department specific data
        item_dataframe = data_frame.loc[data_frame[slice] == item]

        # create the powerpoint for the department
        print('Creating', item, 'ppt.', len(item_dataframe))
        create_pap_pptx(item_dataframe, ppt_deck, item, output_ppt_filename)


if __name__ == "__main__":
    ''' This script pulls two arguments, one for the csv data source and
        the other for the output data file base directory and path'''

    # Pull two arguments - one for the source data and the other for the
    # output file folder name and prefix
    parser = argparse.ArgumentParser()
    parser.add_argument("dataSource", type=str,
                        help="data source")
    parser.add_argument("pptxDestination", type=str,
                        help="destination pptx")
    args = parser.parse_args()

    # get the data from the data file
    print('Creating data frame.')
    pap_data_frame = getDataFrame(args.dataSource)

    # print the PAP ppt
    print('Creating pap pptx.')
    create_pap_pptxs(pap_data_frame, "ppt_template.pptx", args.pptxDestination)
