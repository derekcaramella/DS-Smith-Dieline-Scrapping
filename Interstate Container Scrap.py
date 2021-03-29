# Import necessary modules.
import docx2txt
import pandas as pd
import os

df_list = []  # Prepare empty list for export dataframe
for dieline in os.listdir('Interstate Container Format'):  # Loop through the directory, which analyzes each dieline
    text = docx2txt.process('Interstate Container Format/' + dieline)  # Routes directory & scraps text
    text = text.replace('\n', ',')  # Removes all alt-enters from string
    # Goal: Remove sequence of commas, but leave one comma for splicing
    for index in range(0, len(text)):  # For each index in the string
        if text[index] == ',' and text[index + 1] == ',':  # If the current index is a comma & the next is a comma
            text = text[:index] + ' ' + text[index + 1:]  # Maintain string, but replace with a space to maintain length

    text = text.replace(' ', '')  # Remove all spaces to narrow string
    # Goal: Remove :, instances for string splicing.
    for index in range(0, len(text)):  # For each index in the string
        if text[index] == ',' and text[index - 1] == ':':  # If the current index is a comma & the previous is a colon
            text = text[:index] + ' ' + text[index + 1:]  # Maintain string, but replace with a space to maintain length

    text = text.replace(' ', '')  # Remove all spaces to narrow string
    text = text.replace('	', '')  # Remove all tabs to narrow string
    list_prompts = text.split(',')  # Delimit by comma, this will pair parameter & process value
    design, description, style, board = '', '', '', ''  # Nullify sample for dataframe
    inner_dimension, finished_piece_weight, over_all_blank, combination = '', '', '', ''  # Nullify sample for dataframe
    for instance in list_prompts:  # Within one file, loop through ever instance
        if 'Design#:' in instance:  # If Design#: exists, capture process value
            design = instance.partition('Design#:')[2]  # Bifurcate string & return process value
        if 'Description:' in instance:  # If Description: exists, capture process value
            # Sometimes the description was comma splice, so grab current & next instance
            description = instance.partition('Description:')[2] + str(list_prompts[list_prompts.index(instance) + 1])
        if 'Style:' in instance:  # If Style: exists, capture process value
            style = instance.partition('Style:')[2]  # Bifurcate string & return process value
        if 'Board:' in instance:  # If Board: exists, capture process value
            board = instance.partition('Board:')[2]  # Bifurcate string & return process value
        if 'I.D.:' in instance:  # If I.D.: exists, capture process value
            inner_dimension = instance.partition('I.D.:')[2]  # Bifurcate string & return process value
        if 'Pieceweight-waste:' in instance:  # If Finished Piece Weight: exists, capture process value
            finished_piece_weight = instance.partition('Pieceweight-waste:')[2]
        if 'Blanksize:' in instance:  # If Over-all Blank: exists, capture process value
            over_all_blank = instance.partition('Blanksize:')[2]  # Bifurcate string & return process value
        if 'Combination:' in instance:  # If Combination: exists, capture process value
            combination = instance.partition('Combination:')[2]  # Bifurcate string & return process value
    # Following one file, compile results into one tuple
    sample = (dieline, design, description, style, board, inner_dimension,
              finished_piece_weight, over_all_blank, combination)
    # Append sample into the staging list for dataframe
    df_list.append(sample)
# Generate dataframe from list of tuples
export_df = pd.DataFrame.from_records(df_list, columns=['File', 'Design', 'Description', 'Style', 'Board',
                                                        'Inner Dimension', 'Piece Weight - Waste',
                                                        'Blank Size', 'Combination'])
# Export dataframe
export_df.to_excel('Interstate Container Specifications.xlsx', index=False)
