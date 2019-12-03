'''
createPPTX creates a powerpoint from a given sample powerpoint.
The source will come from a project approval.
'''

# imports go here
from pptx import Presentation
# from pptx.util import Inches
import argparse
import pandas as pd
import os
from datetime import date

# Classes and globals

# Placeholder layout and indices for the Title Slide
TITLE_SLIDE_LAYOUT = 0
TITLE_SLIDE = {'subtitle': 1}

BLANK_SLIDE_LAYOUT = 7

# Placeholder indices for the first approval slide
APP_SLIDE1_LAYOUT = 2
APP_SLIDE1 = {'Key': 27,
              'Sponsor Department': 23,
              'Description': 1,
              'Status': 24,
              'Return Type': 25,
              'ROI': 15,
              'Return': 16,
              'Total Investment': 17,
              'Sponsor': 18,
              'Project Owner': 19,
              'Line of Business': 21,
              'Return Description': 28,
              'Investment Description': 13,
              'Project Manager': 20,
              'States': 22,
              }

# Placeholder indices for the second approval slide
APP_SLIDE2_LAYOUT = 3
APP_SLIDE2 = {'Key': 27,
              'Scope': 14,
              'Scope Exclusions': 15,
              'Systems': 17,
              'Business Area Impacts': 16,
              }

# Data Column Names for simplified mapping
COLMUN_NAMES = {'Issue key': 'Key',
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
                'Custom field (Business Area Impacts)':
                    'Business Area Impacts',
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
    name_df = pd.read_csv("name_mapping.csv")
    name_dict = name_df.to_dict()
    name_dict = {name_df['Username'][ind]: name_df['Full Name'][ind]
                 for ind in name_df.index}

    data_frame["Project Manager"] = data_frame["Project Manager"].map(
                                        name_dict)
    data_frame["Sponsor"] = data_frame["Sponsor"].map(
                                        name_dict)
    return data_frame


def clean_data_frame(data_frame):
    ''' Cleans the data in the data_frame. '''
    # change usernames to Full Names
    data_frame = convert_usernames_to_fullnames(data_frame)
    return data_frame


def getDataFrame(data_file):
    data_frame = pd.read_csv(data_file)
    data_frame.rename(columns=COLMUN_NAMES, inplace=True)
    data_frame = clean_data_frame(data_frame)
    return data_frame


def create_pap_pptx(df, prs, dept, output_filename):
    # Insert title slide with department and status_category such as
    #    project request, inflight or complete.
    for status_category in STATUS_MAPPING:
        statuses = STATUS_MAPPING[status_category]
        status_dataframe_subset = df.loc[df['Status'].isin(statuses)]
        if not status_dataframe_subset.empty:
            title_txt = dept + ' ' + status_category
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
                print(dept, status)
                df_subset = df.loc[(df['Sponsor Department'] == dept) &
                                   (df['Status'] == status)]
                df_subset = df_subset.sort_values(by=['Summary'])
                # Proposed Project Slides
                for index, series in df_subset.iterrows():
                    print('Creating approval slide', index + 1, 'of',
                          len(df_subset.index), '.')
                    title_txt = series["Summary"]
                    populateSlideFromSeries(createSlide(prs,
                                                        APP_SLIDE1_LAYOUT),
                                            title_txt, APP_SLIDE1, series)
                    populateSlideFromSeries(createSlide(prs,
                                                        APP_SLIDE2_LAYOUT),
                                            title_txt, APP_SLIDE2, series)

        prs.save(output_filename)


def create_pap_pptxs(dataframe, ppt_template, file_prefix):
    '''Create a project approval ppt from a dataframe and ppt template.
        A pap ppt deck will be created for each dept with the output prefix.'''

    # make a folder for the ppts
    ppt_folder_path = file_prefix.replace('.pptx', '')
    if not os.path.exists(ppt_folder_path):
        os.makedirs(ppt_folder_path)

    # create a base path and filename for the ppts
    ppt_base_file_path = os.path.join(
                        ppt_folder_path,
                        os.path.basename(file_prefix) + '.pptx')

    # create a ppt deck for each department
    for dept in dataframe['Sponsor Department'].unique():
        # build a ppt deck for each of the departments
        ppt_deck = Presentation(ppt_template)

        # create the output filename using the file_prefix and department
        output_ppt_filename = ppt_base_file_path.replace(
                                '.pptx', '_' + str(dept).replace(' ', '_') +
                                '.pptx')
        # get the department specific data
        dept_dataframe = dataframe.loc[dataframe['Sponsor Department'] == dept]

        # create the powerpoint for the department
        print('Creating', dept, 'ppt.')
        create_pap_pptx(dept_dataframe, ppt_deck, dept, output_ppt_filename)


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
