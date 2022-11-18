import string
from datetime import date
import json
import docx
import csv
import sys

def fill_template(parameter_dict: dict):

    # Open file 
    
    # Create copy of file

    # fill copy with parameters

    # Save Copy
    pass

def main():

    with open("config.json", "r") as json_file:
        config = json.load(json_file)

    # If "-b" (batch) then enter loop over each row in batches.csv
    if sys.argv[1] == "-b":
    
        with open("batches.csv", "r") as csv_file:
            csv_reader = csv.reader(csv_file)

            for row in csv_reader:

                parameters = {"Company Name": row[0],
                              "Letter Type": row[1],
                              "Recuiter Name": row[2],
                              "Position Title": row[3],
                              "Current Date": f"{date.today().month}/{date.today().day}/{date.today().year}"}

                fill_template(parameters)

    # If "-s" (single) then fill template by hard-coded dictionary
    elif sys.argv[1] == "-s":

        parameters = {"Company Name": "TEST_COMPANY_NAME",
                        "Letter Type": "tempate_1",
                        "Recuiter Name": "RECRUITER_NAME",
                        "Position Title": "POSITION_TITLE",
                        "Current Date": f"{date.today().month}/{date.today().day}/{date.today().year}"}

        fill_template(parameters)

    # Unknown parameter for Cover Letter generation
    else:
        print("\nError: unknown execution param, either \"-s\" or \"-b\"...\n")

if __name__ == "__main__":
    
    main()
