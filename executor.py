import os, argparse
from xlrd import open_workbook
from utils import Utils
from selenium import webdriver


TEST_CASE_PATH = '\\test_cases'
objectRepo_dict = {}   # a dictionary of all the xpath values from repository sheet.
variable_dict = {}     # a dictionary of all the variables from the test_data sheet.
testcases_dict = {}     # a dictionary of all the executatble test cases.


class Executor:

    def setup(self, config_params):
        """
        Validate the file name and sheet name provided by the user.
        :param config_params: Command line parameters
        :return:
        """
        file_path = os.path.join(os.getcwd()+TEST_CASE_PATH, config_params['file_name']+'.xlsx')
        sheet_name = config_params['sheet_name']

        if not os.path.isfile(file_path): # check if the file exists or not
            raise ValueError(file_path + ' is not valid.') # throw some error over here.

        wb = open_workbook(file_path)

        # check if Test case sheet is present or not.
        if  sheet_name not in [sheet.name for sheet in wb.sheets()]:
            raise ValueError(sheet_name + 'not found in the xls file.')

        self.get_object_repository(wb)  # store object repository in dict
        self.get_variables(wb)  # store the variables in a dict
        self.get_executable_data(wb, sheet_name, config_params['group_names'].split(','))
        self.execute_testcases()

    def get_object_repository(self, wb):
        """
        Fetch the locators from Repository sheet and store them in a dictionary
        :param wb: workbook object
        :return:
        """
        repo_sheet = wb.sheet_by_name('repository')
        headers = self.get_sheet_headers(repo_sheet)  # Get the headers values of the 'Repository' Sheet.
        for row in range(repo_sheet.nrows):
            if row != 0:
                repo = {}
                repo['Locator_Value'] = repo_sheet.cell(row, headers['Locator_Value']).value
                repo['Locator_Type'] = repo_sheet.cell(row, headers['Locator_Type']).value
                objectRepo_dict[repo_sheet.cell(row, headers['Name']).value] = repo
        #print(objectRepo_dict)

    def get_variables(self, wb):
        """
        Fetch the variables from 'Test_Data' sheet and store them in a dictionary
        :param wb:
        :return:
        """
        headers = {}
        var_sheet = wb.sheet_by_name('test_data')
        for row in range(var_sheet.nrows):
            var_obj = {}
            for col in range(var_sheet.ncols):
                data = var_sheet.cell(row, col).value
                if row == 0:
                    headers[col] = data
                else:
                    var_obj[headers[col]] = data
            if row != 0:
                variable_dict[var_sheet.cell(row, 0).value] = var_obj
        #print(variable_dict)

    def get_executable_data(self, wb, sheet_name, group_names):
        """
        Fetch the executed test cases from the Test Case file.
        :param wb: TestCase workbook being used.
        :param sheet_name: TestCase sheet name
        :param group_names: Group names that we want to run
        :return: N/A
        """
        tc_sheet = wb.sheet_by_name(sheet_name)
        headers = self.get_sheet_headers(tc_sheet)
        test_case_id = tc_sheet.cell(1, headers['TestCaseID']).value

        row_data_list = []

        for row in range(tc_sheet.nrows):
            if row != 0:
                if test_case_id != tc_sheet.cell(row, headers['TestCaseID']).value:
                    if len(row_data_list)>0:
                        testcases_dict[test_case_id] = row_data_list
                    test_case_id = tc_sheet.cell(row, headers['TestCaseID']).value
                    row_data_list = []

                action_step = tc_sheet.cell(row, headers['Action']).value
                name = tc_sheet.cell(row, headers['Name']).value
                groups = tc_sheet.cell(row, headers['Groups']).value
                row_data = {}
                if action_step[0:4] == 'cus_':
                    row_data['Name'] = name
                    row_data['Groups'] = groups
                    row_data['Action'] = action_step
                    row_data['Locator'] = tc_sheet.cell(row, headers['Locator']).value
                    row_data['I/O Values'] = tc_sheet.cell(row, headers['I/O Values']).value
                    row_data_list.append(row_data)
                else:
                    action_step_data = self.get_action_steps(wb, name, groups, action_step)
                    for step_data in action_step_data:
                        row_data_list.append(step_data)

                if row == sum(1 for x in tc_sheet.get_rows()) - 1:
                    if len(row_data_list) > 0:
                        testcases_dict[test_case_id] = row_data_list
                    test_case_id = tc_sheet.cell(row, headers['TestCaseID']).value
        #print(testcases_dict)

    def get_action_steps(self, wb, name, groups, action_step):
        """
        Fetch the common TestCases from the 'action_step' sheet.
        :param wb: workbook being used.
        :param name: TestCase name
        :param groups: Group names that we want to run
        :param action_step: The action step to be looked into in the 'action_steps'
        :return: N/A
        """
        step_matched= False
        action_step_sheet = wb.sheet_by_name('action_steps')
        headers = self.get_sheet_headers(action_step_sheet)
        step = action_step
        row_data_list = []

        for row in range(action_step_sheet.nrows):
            if row != 0:
               if action_step_sheet.cell(row, headers['Steps']).value == step:
                   step_matched = True
                   row_data = {}
                   row_data['Name'] = action_step_sheet.cell(row, headers['Name']).value
                   row_data['Groups'] = groups
                   row_data['Action'] = action_step_sheet.cell(row, headers['Action']).value
                   row_data['Locator'] = action_step_sheet.cell(row, headers['Locator']).value
                   row_data['I/O Values'] = action_step_sheet.cell(row, headers['I/O Values']).value
                   row_data_list.append(row_data)
                   if row == sum(1 for x in action_step_sheet.get_rows()) - 1:  # if its last row of the sheet.
                       return row_data_list
               else:
                   if step_matched is True:   # if there is no more matches for the action step.
                       return row_data_list

        if step_matched is False:
            raise ValueError('Step -' + action_step + ' not found in aciton_steps sheet')


    def get_sheet_headers(self, sheet):
        """
        Fetch the headers from a sheet and store them to a dictionary
        :param sheet: The Sheet name to get Headers data from.
        :return:
        """
        headers = {}
        for row in range(sheet.nrows):
            for col in range(sheet.ncols):
                data = sheet.cell(row, col).value
                if row == 0:
                    headers[data] = col
                else:
                    break
        return headers

    def execute_testcases(self):
        """
        Execute the test cases
        :return: N/A
        """
        for test_case in testcases_dict.keys():
            print('Starting Execution of - ' + test_case)
            util = Utils()
            for index, actions in enumerate(testcases_dict[test_case]):
                print("Running Step: " + str(index) + ' ---> '+ actions['Name'])
                getattr(util, actions['Action'][4:])(actions, variable_dict[test_case], objectRepo_dict)
            util.tear_down()  # quit the browser after each TestCase


if __name__ == '__main__':
    executor = Executor()
    parser = argparse.ArgumentParser()
    parser.add_argument("-f", "--file", required="True", help="TestCase file name is required.")
    parser.add_argument("-s", "--sheet", required="True", help="TestCase sheet in the TestCase file.")
    parser.add_argument("-g", "--groupNames", required="True", help="Group names to be executed.")
    args = parser.parse_args()

    # storing config data
    global_config = {}
    global_config['file_name'] = args.file
    global_config['sheet_name'] = args.sheet
    global_config['group_names'] = args.groupNames

    executor.setup(global_config)  # Automation test Setup


