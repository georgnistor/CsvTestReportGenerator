from data_structure_csv import *
import argparse
import os

# create parser
parser = argparse.ArgumentParser()
# add arguments to the parser
parser.add_argument("lava_file")
parser.add_argument("git_repo")
# parse the arguments
args = parser.parse_args()
ltp___reportFile = args.lava_file
path_ltp_git = ((args.git_repo.split('/'))[-1]).split('.')[0]  # extract the root folder of the git repository

if os.path.isdir(path_ltp_git):
    if os.name == 'nt':
        os.system('rmdir /s /q ' + path_ltp_git)
    else:
        os.system('rm -r ' + path_ltp_git)
    print('ltp git repository is present -> it will be deleted')

os.system('git clone ' + args.git_repo)

Generator.git_runtest_extract_data(path_ltp_git)
Generator.file_parser_ltp(ltp___reportFile)

"""optional (output from the data_structure)"""
#Generator.list_test_cases()

Generator.format_excel_sheet()


