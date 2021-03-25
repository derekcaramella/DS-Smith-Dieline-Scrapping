# Import necessary modules.
import docx2txt
import os

for dieline in os.listdir('DS Smith Dielines'):  # Loop through the directory, which analyzes each dieline
    text = docx2txt.process('DS Smith Dielines/' + dieline)  # Routes directory & scraps text
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

    print(text)  # Return comprehensive string for splicing
