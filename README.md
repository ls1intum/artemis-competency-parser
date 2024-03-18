# Artemis Competency Parser

## About
The purpose of this repository is to transform the core topics defined in the Computer Science Curricula 2023 (CS2023) into Artemis-compatible standardized competencies.
To achieve this, it defines a semi-automatic pipeline that takes an excel workbook containing these core topics and converts them into a .json file of knowledge areas and competencies that can be imported into Artemis.

## Pipeline: Steps for execution
1. Copy Word Tables -> Excel Sheets (**manual**)
    1. Copy the whole table for a Knowledge Area (KA) from the *Core Topics Table* chapter in the CS2023
    2. Paste it into a new sheet in the *CS2023_Knowledge_Areas.xlsm* file
    3. Name the sheet like the abbreviation of the KA
    4. Repeat for all KAs
2. Mark new objects (VBA)
	1. Open the excel file and press Alt + F11 to go to the VBA view
    2. Run the *markNewOnAllWorksheets* routine
3. .csv -> .json (python)
	1. run `pip install -r requirements.txt` to install all requirements
    2. run `python competency_parser.py`
4. error detection/handling (**manual**)
	1. Most likely not all of the competencies will be immediately correct. They have to be fixed manually.
    2. Open the file *data/3_error_competencies.json*
    3. See what problems the comeptencies have by their *error* field
    4. Fix those errors (refer to the chapter **Competency error types** below)
	5. (Optional) if you want to check the status run `python competency_parser.py --verify`
		Note: this also moves corrected competencies to the *data/3_correct_competencies.json*) file
5. generating an import ready .json (python)
	1. When done correcting, execute `python competency_parser.py --finalize`
    2. If you fixed all competencies, the script runs through and saves the competencies to *competencies_for_import.json*
    3. If not, repeat step 4
	
## Python parser 
This chapter contains some detailed information on the python parser `competency_parser.py`

### Competency error types
- DUPLICATE_TITLE: 
    - Cause: multiple competencies have the same title. 
    - Solution: Change (at least) one of them to something different.
- TITLE_TOO_LONG: 
    - Cause: The title is too long. 
    - Solution: Shorten it to at most 255 characters (defined in `MAX_TITLE_LENGTH`)
- DESCRIPTION_TOO_LONG: 
    - Cause: The description is too long. 
    - Solution: Shorten it to at most 2000 characters (defined in `MAX_DESCRIPTION_LENGTH`)
- WRONG_TAXONOMY: 
    - Cause: The competency has a taxonomy that is neither part of the "Skill Levels" in CS2023 nor in the CompetencyTaxonomies of Artemis
    - Solution: change it to one of the allowed values (see `ALLOWED_TAXONOMIES`): 
        - "REMEMBER", "UNDERSTAND", "APPLY", "ANALYZE", "EVALUATE", "CREATE" (Artemis)
        - "Explain", "Apply", "Evaluate", "Develop" (CS2023, will be mapped to Artemis taxonomies later)
- WRONG_KNOWLEDGE_AREA:
    - Cause: The competency has a knowledge area that is unknown
    - Solution: change it to one of the allowed values in the `KNOWLEDGE_AREA_MAPPING`
    - Note: this should **not** happen, except for human error. If new knowledge areas are introduced this should be caught by the first step of the python script

### TODO
TODO: add settings etc. (Especially do_backup)
	
### Detailed steps
The parser consists of the following five steps:
1. Convert excel to raw competencies: 
	- Takes the excel workbook, verifies all columns are present for every sheet
	- Saves the competencies raw (i.e. exactly as they are in the excel sheet)
2. Convert to clean competencies
	- Takes raw competencies as input
	- Removes all unneeded information from the competencies
	- Re-names columns
3. Mark errors
	- Takes clean competencies as input
	- Marks all errors in the competencies (see the chapter **Competency error types** above)
	- Saves all error competencies and all correct competencies seperately
4. Verify Competencies
	- Runs the error marking step again for all error competencies
	- Saves any that were corrected to the correct competencies
5. Convert to Artemis
	- Can only be executed if all error competencies have been fixed
	- Converts all correct competencies to a format that allows it to import them into Artemis
	- For now this is a list of knowledge areas containing these competencies
	
### Python parser: Command line arguments
- `--verify`: Runs the verification step 
- `--finalize`: Runs the verification step and generates the artemis import file (if no error competecies exist)
- `--step`: Runs a specific step. Possible values: excel_to_raw, raw_to_clean, mark_errors, verify, convert_to_artemis, all