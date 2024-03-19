import pandas as pd
import argparse
import json
import sys
import os 

### FILE CONFIGURATION
# get overwritten in main
BASE_DIRECTORY = ""
BACKUP_DIRECTORY = ""
# all paths are relative to the BASE_DIRECTORY
DATA_DIRECTORY = "data"
PREVIOUS_RUN_DIRECTORY = os.path.join(DATA_DIRECTORY, "previous_runs")

EXCEL_INPUT_FILE = "CS2023_Knowledge_Areas.xlsm"
RAW_COMPETENCIES_FILE = os.path.join(DATA_DIRECTORY, "1_raw_competencies.json")
CLEAN_COMPETENCIES_FILE = os.path.join(DATA_DIRECTORY, "2_clean_competencies.json")
CORRECT_COMPETENCIES_FILE = os.path.join(DATA_DIRECTORY, "3_correct_competencies.json")
ERROR_COMPETENCIES_FILE = os.path.join(DATA_DIRECTORY, "3_error_competencies.json")
FINAL_COMPETENCIES_FILE = "competencies_for_import.json"
RUN_INFO_FILE = "run_info.json"

### EXECUTION CONFIGURATION
# if the competency data should be backed up for each run under PREVIOUS_RUN_DIRECTORY
DO_BACKUPS = True

# mapping knowledge area abbreviations -> full title
# this is needed as we cannot have the full title for excel worksheets
KNOWLEDGE_AREA_MAPPING = {
    "AI": "Artificial Intelligence",
    "AL": "Algorithmic Foundations",
    "AR": "Architecture and Organization",
    "DM": "Data Management",
    "FPL": "Foundations of Programming Languages",
    "GIT": "Graphics and Interactive Techniques",
    "HCI": "Human-Computer Interaction",
    "MSF": "Mathematical and Statistical Foundations",
    "NC": "Networking and Communication",
    "OS": "Operating Systems",
    "PDC": "Parallel and Distributed Computing",
    "SDF": "Software Development Fundamentals",
    "SE": "Software Engineering",
    "SEC": "Security",
    "SEP": "Society, Ethics, and the Profession",
    "SF": "Systems Fundamentals",
    "SPD": "Specialized Platform Development",
}

# raw competency/excel file columns
TITLE_COLUMN_RAW = "KU"
DESCRIPTION_COLUMN_RAW = "Topic"
TAXONOMY_COLUMN_RAW = "Skill Level"
NEW_OBJECT_COLUMN_RAW = "new_in_next_row"
KNOWLEDGE_AREA_COLUMN_RAW = "KA"
# these are the columns we expect in the excel files. Note that KA is not present as is gets added later.
EXCEL_COLUMNS = [TITLE_COLUMN_RAW, DESCRIPTION_COLUMN_RAW, TAXONOMY_COLUMN_RAW, NEW_OBJECT_COLUMN_RAW]

# mapping from raw competencies -> clean/correct/error competencies
TITLE_COLUMN = "title"
DESCRIPTION_COLUMN = "description"
TAXONOMY_COLUMN = "taxonomy"
KNOWLEDGE_AREA_COLUMN = "knowledgeArea"
ERROR_COLUMN = "error"
COLUMN_MAPPING_RAW_CLEAN = {
    TITLE_COLUMN_RAW: TITLE_COLUMN,
    DESCRIPTION_COLUMN_RAW: DESCRIPTION_COLUMN,
    TAXONOMY_COLUMN_RAW: TAXONOMY_COLUMN,
    KNOWLEDGE_AREA_COLUMN_RAW: KNOWLEDGE_AREA_COLUMN,
}

# mapping CS2023 taxonomies -> artemis taxonomies
TAXONOMY_MAPPING = {
    "Explain": "UNDERSTAND",
    "Apply": "APPLY",
    "Evaluate": "EVALUATE",
    "Develop": "CREATE"
}
ARTEMIS_TAXONOMIES = ["REMEMBER", "UNDERSTAND", "APPLY", "ANALYZE", "EVALUATE", "CREATE"]

# Settings for verification
ALLOWED_TAXONOMIES = list(TAXONOMY_MAPPING.keys()) + ARTEMIS_TAXONOMIES
ALLOWED_KNOWLEDGE_AREAS = list(KNOWLEDGE_AREA_MAPPING.values()) + list(KNOWLEDGE_AREA_MAPPING.keys())

MAX_TITLE_LENGTH = 255
MAX_DESCRIPTION_LENGTH = 2000

### CLI ARGUMENT CONFIGURATION
EXCEL_TO_RAW = "excel_to_raw"
RAW_TO_CLEAN = "raw_to_clean"
MARK_ERRORS = "mark_errors"
VERIFY = "verify"
CONVERT_TO_ARTEMIS = "convert_to_artemis"
ALL_STEPS = [EXCEL_TO_RAW, RAW_TO_CLEAN, MARK_ERRORS, VERIFY, CONVERT_TO_ARTEMIS]

### FILE UTILITY FUNCTIONS

def write_to_file(content, output_file):
    content_json = json.dumps(content, indent=4)
    with open(os.path.join(BASE_DIRECTORY, output_file), "w") as outfile:
        outfile.write(content_json)
        outfile.close()

def write_to_file_and_backup(content, output_file):
    write_to_file(content, output_file)
    if DO_BACKUPS:
        file_name = os.path.basename(output_file)
        backup_file = os.path.join(BACKUP_DIRECTORY, "backup_" + file_name)
        write_to_file(content, backup_file)

def load_from_file(input_file):
    if not is_file_usable(input_file):
        print(f"File {input_file} is empty or does not exist!")
        sys.exit(0)
    file = open(os.path.join(BASE_DIRECTORY, input_file))
    content = json.load(file)
    file.close()
    return content

def is_file_usable(file):
    full_file_path = os.path.join(BASE_DIRECTORY, file)
    if not does_file_exist(file):
        return False
    if os.stat(full_file_path).st_size == 0:
        return False
    return True

def does_file_exist(file):
    full_file_path = os.path.join(BASE_DIRECTORY, file)
    if os.path.isfile(full_file_path):
        return True
    return False

# UTILITY FUNCTIONS FOR THE EXECUTION

def get_and_increase_run_count():
    run_info = load_from_file(RUN_INFO_FILE)
    run_info["number"] += 1
    write_to_file(run_info, RUN_INFO_FILE)
    return run_info["number"]


def parse_arguments(argv):
    parser = argparse.ArgumentParser(
            description="Parse an excel workbook with CS2023 knowledge units into Artemis competencies",
            formatter_class=argparse.RawTextHelpFormatter,
        )
    parser.add_argument(
            "--finalize", "-F",
            action="store_true",
            help='''Executes the verification step and creates the Artemis import file'''
        )
    parser.add_argument(
            "--verify", "-V",
            action="store_true",
            help='''Executes the verification step'''
        )
    # add a possibility to execute all steps
    steps = ALL_STEPS + ["all"]
    parser.add_argument(
            "--step", "-S",
            choices=steps, 
            action="store",
            help='''Only use this argument if you want to execute an exact step and know what you are doing!'''
        )
    
    args = parser.parse_args(argv)
    exclusive_args = ["finalize", "verify", "step"]
    exclusive_arg_count = sum([(1 if vars(args)[arg] else 0) for arg in exclusive_args])

    if exclusive_arg_count > 1:
        print(f"The arguments [{', '.join(exclusive_args)}] exclude each other, please only use one of them!")
        sys.exit(0)

    if args.verify:
        return [VERIFY]
    if args.finalize:
        return [VERIFY, CONVERT_TO_ARTEMIS]
    if args.step is None or args.step == "all":
        return ALL_STEPS
    return [args.step]

# STEP 1 FUNCTIONS

# if any data is missing, take it from the previous competency
# this happens if a competency has multiple versions in the table, e.g. different taxonomies but only one title -> later versions have no title
def correct_values(competency, last_competency, knowledge_area):
    for key in competency:
        competency[key] = competency[key].strip()
        if competency[key] == "":
            competency[key] = last_competency[key]
    # append the current knowledge area to the competency
    competency["KA"] = knowledge_area
    return competency

def create_empty_competency(columns):
    # save all columns except for NEW_OBJECT_COLUMN_RAW
    competency = dict.fromkeys(columns, "")
    competency.pop(NEW_OBJECT_COLUMN_RAW, None)
    return competency

def get_competencies_for_sheet(sheet_name, dataframe):
    knowledge_area = sheet_name
    columns = list(dataframe.columns)

    competencies = []
    currenct_competency = create_empty_competency(columns)
    last_competency = create_empty_competency(columns)

    for _, row in dataframe.iterrows():
        # add values to current competency (unless they are null)
        for key in currenct_competency:
            if not pd.isnull(row[key]):
                currenct_competency[key] += str(row[key]) + "\n"
        # if the next row is a new object, save to list and clear current competency
        if row[NEW_OBJECT_COLUMN_RAW] == NEW_OBJECT_COLUMN_RAW:
            currenct_competency = correct_values(currenct_competency, last_competency, knowledge_area)
            competencies.append(currenct_competency)
            last_competency = currenct_competency
            currenct_competency = create_empty_competency(columns)
    
    return competencies

def are_sheets_ok(sheets_as_dataframe_dict):
    print("Verifying the excel sheets...")
    sheets_ok = True
    for sheet_name in sheets_as_dataframe_dict:
        dataframe_columns = list(sheets_as_dataframe_dict[sheet_name].columns)
        missing_columns = []
        for col in EXCEL_COLUMNS:
            if col not in dataframe_columns:
                missing_columns.append(col)
        if missing_columns:
            print(f"The sheet \"{sheet_name}\" is missing columns: {missing_columns}")
            sheets_ok = False
        if not sheet_name in ALLOWED_KNOWLEDGE_AREAS:
            print(f"The sheet \"{sheet_name}\" has a name that is not part of the KNOWLEDGE_AREA_MAPPING. Please add it!")
            sheets_ok = False
    return sheets_ok

# STEP 2 FUNCTIONS

def convert_to_clean(raw_competency):
    clean_competency = {}
    # remove unused columns and rename columns
    for key in COLUMN_MAPPING_RAW_CLEAN:
        mapped_key = COLUMN_MAPPING_RAW_CLEAN[key]
        clean_competency[mapped_key] = raw_competency[key]
        # replace nbsp as those get created by formatting errors in word
        clean_competency[mapped_key] = clean_competency[mapped_key].replace(u'\u00a0', u'')
    return clean_competency

# STEP 3/4 FUNCTIONS

def mark_errors(competencies):
    # mark errors for all competencies and save them to files
    error_competencies = []
    clean_competencies = []
    
    # check for duplicate titles
    previous_title = ""
    for i in range(len(competencies)):
        current_title = competencies[i][TITLE_COLUMN]
        if current_title == previous_title:
            competencies[i][ERROR_COLUMN] = "DUPLICATE_TITLE"
            competencies[i-1][ERROR_COLUMN] = "DUPLICATE_TITLE"
        previous_title = current_title

    # check for errors in the fields
    for competency in competencies:
        competency = mark_field_errors(competency)
        if ERROR_COLUMN in competency:
            error_competencies.append(competency)
        else:
            clean_competencies.append(competency)
    return (error_competencies, clean_competencies)

def mark_field_errors(competency):
    errors = []

    # title errors
    title = competency[TITLE_COLUMN]
    if "\n" in title:
        errors.append("MULTIPLE_TITLES")
    if len(title) > MAX_TITLE_LENGTH:
        errors.append("TITLE_TOO_LONG")

    # description errors
    description = competency[DESCRIPTION_COLUMN]
    if len(description) > MAX_DESCRIPTION_LENGTH:
        errors.append("DESCRIPTION_TOO_LONG")

    # taxonomy errors
    taxonomy = competency[TAXONOMY_COLUMN]

    if not taxonomy in ALLOWED_TAXONOMIES:
        errors.append("WRONG_TAXONOMY")

    # knowledge area errors
    knowledge_area = competency[KNOWLEDGE_AREA_COLUMN]
    if not knowledge_area in ALLOWED_KNOWLEDGE_AREAS:
        errors.append("WRONG_KNOWLEDGE_AREA")

    # if errors give the competency an extra field
    if errors:
        if not ERROR_COLUMN in competency:
            competency[ERROR_COLUMN] = ""
        competency[ERROR_COLUMN] += "," + ','.join(errors)
    return competency

# STEP 5 FUNCTIONS

def convert_to_artemis_format(competency):
    # copy competency to not change the original
    competency_copy = competency.copy()
    taxonomy = competency_copy[TAXONOMY_COLUMN]
    # convert taxonomy to artemis format
    if taxonomy in TAXONOMY_MAPPING:
        competency_copy[TAXONOMY_COLUMN] = TAXONOMY_MAPPING[taxonomy]

    # remove knowledge area from competency
    competency_copy.pop(KNOWLEDGE_AREA_COLUMN, None)

    return competency_copy

# FUNCTIONS FOR EACH EXECUTION STEP

def s1_convert_excel_to_raw_competencies():
    print("Starting conversion excel -> raw competencies\n=====")
    if does_file_exist(RAW_COMPETENCIES_FILE):
        print(f"{RAW_COMPETENCIES_FILE} already exists! If you want to execute this step, please move it")
        print("=====\nCanceled conversion raw -> clean competencies\n")
        sys.exit(0)

    sheets_as_dataframe_dict = pd.read_excel(os.path.join(BASE_DIRECTORY,EXCEL_INPUT_FILE), sheet_name=None)
    print(f"Loaded {len(sheets_as_dataframe_dict)} sheets: {list(sheets_as_dataframe_dict.keys())}")
    raw_competencies = []

    if not are_sheets_ok(sheets_as_dataframe_dict):
        print("=====\nCanceled conversion excel -> raw competencies")
        sys.exit(0)
    else:
        print("All sheets are ok. Continuing...")

    for sheet_name in sheets_as_dataframe_dict:
        competencies_in_sheet = get_competencies_for_sheet(sheet_name, sheets_as_dataframe_dict[sheet_name])
        raw_competencies += competencies_in_sheet
    write_to_file_and_backup(raw_competencies, RAW_COMPETENCIES_FILE)
    print(f"Saved {len(raw_competencies)} (raw) competencies to {RAW_COMPETENCIES_FILE}")
    print("=====\nFinished conversion excel -> raw competencies\n")

def s2_convert_to_clean_competencies():
    print("Starting conversion raw -> clean competencies\n=====")
    if does_file_exist(CLEAN_COMPETENCIES_FILE):
        print(f"{CLEAN_COMPETENCIES_FILE} already exists! If you want to execute this step, please move it")
        print("=====\nCanceled conversion raw -> clean competencies\n")
        sys.exit(0)

    raw_competencies = load_from_file(RAW_COMPETENCIES_FILE)
    competencies = []

    for competency in raw_competencies:
        clean_competency = convert_to_clean(competency)
        competencies.append(clean_competency)
    write_to_file_and_backup(competencies, CLEAN_COMPETENCIES_FILE)
    print(f"Saved {len(competencies)} (clean) competencies to {CLEAN_COMPETENCIES_FILE}")
    print("=====\nFinished conversion raw -> clean competencies\n")

def s3_mark_errors():
    print("Starting error marking\n=====")
    if does_file_exist(CORRECT_COMPETENCIES_FILE) or does_file_exist(ERROR_COMPETENCIES_FILE):
        print(f"{CORRECT_COMPETENCIES_FILE} and/or {ERROR_COMPETENCIES_FILE} already exist! If you want to execute this step, please move them")
        print("=====\nCanceled marking errors\n")
        sys.exit(0)

    competencies = load_from_file(CLEAN_COMPETENCIES_FILE)
    print(f"Loaded {len(competencies)} competencies")

    error_competencies, correct_competencies = mark_errors(competencies)

    write_to_file_and_backup(error_competencies, ERROR_COMPETENCIES_FILE)
    write_to_file_and_backup(correct_competencies, CORRECT_COMPETENCIES_FILE)
    print(f"Saved {len(correct_competencies)} (correct) competencies to {CORRECT_COMPETENCIES_FILE}")
    print(f"Saved {len(error_competencies)} (error) competencies to {ERROR_COMPETENCIES_FILE}")
    print("=====\nFinished marking errors\n")

def s4_verify_competencies():
    print("Starting verification\n=====")
    if not is_file_usable(ERROR_COMPETENCIES_FILE):
        print(f"Error competencies file {ERROR_COMPETENCIES_FILE} is empty or does not exist!")
        print("=====\nCanceld verification\n")
        return

    old_error_competencies = load_from_file(ERROR_COMPETENCIES_FILE)
    old_correct_competencies = load_from_file(CORRECT_COMPETENCIES_FILE)
    print(f"Loaded {len(old_correct_competencies)} (correct) and {len(old_error_competencies)} (error) competencies")

    # remove error column
    for competency in old_error_competencies:
        competency.pop(ERROR_COLUMN, None)

    # mark errors anew
    error_competencies, correct_competencies = mark_errors(old_error_competencies)
    print(f"Corrected {len(correct_competencies)} competencies, {len(error_competencies)} competencies with errors remain")

    # append new correct competencies to the existing ones
    correct_competencies += old_correct_competencies

    write_to_file_and_backup(error_competencies, ERROR_COMPETENCIES_FILE)
    write_to_file_and_backup(correct_competencies, CORRECT_COMPETENCIES_FILE)
    print(f"Saved {len(old_correct_competencies)} (correct) competencies to {CORRECT_COMPETENCIES_FILE}")
    print(f"Saved {len(error_competencies)} (error) competencies to {ERROR_COMPETENCIES_FILE}")
    print("=====\nFinished verification\n")

def s5_convert_to_artemis():
    print("Starting conversion to Artemis import file\n=====")
    if does_file_exist(FINAL_COMPETENCIES_FILE):
        print(f"{FINAL_COMPETENCIES_FILE} already exists! If you want to execute this step, please move it")
        print("=====\nCanceled conversion to Artemis import file\n")
        sys.exit(0)
    if is_file_usable(ERROR_COMPETENCIES_FILE):
        error_competencies = load_from_file(ERROR_COMPETENCIES_FILE)
        if error_competencies:
            print(f"ERROR: {len(error_competencies)} competencies with errors still exist.")
            print(f"Please fix all errors in the file {ERROR_COMPETENCIES_FILE} and execute the program again with:")
            print("--finalize: to verify all errors are fixed and to create the Artemis import file OR")
            print("--verify: to see how many errors have been fixed (the import file still has to be created afterwards with --verify)")
            print(f"If you wish to brute-force the creation of the import file, delete the file {ERROR_COMPETENCIES_FILE}")
            print("=====\nCanceled conversion to Artemis import file\n")
            sys.exit(0)
    else:
        print(f"No error competencies found for file {ERROR_COMPETENCIES_FILE}. Continuing...")

    competencies = load_from_file(CORRECT_COMPETENCIES_FILE)
    print(f"Loaded {len(competencies)} competencies from {CORRECT_COMPETENCIES_FILE}")
    num_competencies = 0

    knowledge_area_keys = list(KNOWLEDGE_AREA_MAPPING.keys())
    knowledge_areas = []
    for ka_key in knowledge_area_keys:
        ka_title = KNOWLEDGE_AREA_MAPPING[ka_key]
        knowledge_area = {
            "title": ka_title,
            "description": "",
            "competencies": []
        }
        for competency in competencies:
            competency_knowledge_area = competency[KNOWLEDGE_AREA_COLUMN]
            # accept either the abbreviation or full knowledge area name
            if competency_knowledge_area == ka_key or competency_knowledge_area == ka_title:
                competency_artemis_format = convert_to_artemis_format(competency)
                knowledge_area["competencies"].append(competency_artemis_format)

        print(f"Saving knowledge area \"{ka_title}\" with {len(knowledge_area["competencies"])} competencies")
        knowledge_areas.append(knowledge_area)
        num_competencies += len(knowledge_area["competencies"])

    write_to_file_and_backup(knowledge_areas, FINAL_COMPETENCIES_FILE)

    print(f"Saved {len(knowledge_areas)} knowledge areas with {num_competencies} competencies to {FINAL_COMPETENCIES_FILE}")
    print("=====\nFinished conversion to Artemis import file\n")

# MAIN FUNCTION

def main(argv):
    # gets which steps should be executed from the provided arguments
    steps_to_execute = parse_arguments(argv)

    # get the directory of main.py to resolve all other paths 
    global BASE_DIRECTORY 
    BASE_DIRECTORY = os.path.dirname(os.path.realpath(__file__))

    # increase run number and create a directory that backs up data for this run
    if DO_BACKUPS:
        run_number = get_and_increase_run_count()
        global BACKUP_DIRECTORY 
        BACKUP_DIRECTORY = os.path.join(BASE_DIRECTORY, PREVIOUS_RUN_DIRECTORY, str(run_number))
        os.mkdir(BACKUP_DIRECTORY)

    print("")

    if EXCEL_TO_RAW in steps_to_execute:
        s1_convert_excel_to_raw_competencies()
    if RAW_TO_CLEAN in steps_to_execute:
        s2_convert_to_clean_competencies()
    if MARK_ERRORS in steps_to_execute:
        s3_mark_errors()
    if VERIFY in steps_to_execute:
        s4_verify_competencies()
    if CONVERT_TO_ARTEMIS in steps_to_execute:
        s5_convert_to_artemis()


if __name__ == "__main__":
    main(sys.argv[1:])