from selenium import webdriver
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By


class Utils:

    def __init__(self):
        self.driver = webdriver.Firefox()
        self.driver.maximize_window()

    def tear_down(self):
        self.driver.quit()

    def go_to_url(self, parameters, variables, objectRepo):
        """
        Navigate to a URL
        :param parameters: Dictionary of parameters to be used with this test case.
        :param variables: Dictionary of variables stored form 'test_data' sheet
        :param objectRepo: Dictionary of element locators from 'repository' sheet.
        :return: N/A
        """
        self.driver.get(variables[parameters['I/O Values']])

    def click(self, parameters, variables, objectRepo):
        """
        Click on an element.
        :param parameters: Dictionary of parameters to be used with this test case.
        :param variables: Dictionary of variables stored form 'test_data' sheet
        :param objectRepo: Dictionary of element locators from 'repository' sheet.
        :return: N/A
        """
        element = self.find_element(parameters, objectRepo)
        element.click()

    def enter_text(self, parameters, variables, objectRepo):
        """
        Enter text in a input box.
        :param parameters: Dictionary of parameters to be used with this test case.
        :param variables: Dictionary of variables stored form 'test_data' sheet
        :param objectRepo: Dictionary of element locators from 'repository' sheet.
        :return:
        """
        element = self.find_element(parameters, objectRepo)
        element.send_keys(variables[parameters['I/O Values']])

    def assert_not_present(self, parameters, variables, objectRepo):
        """
        Assert that an element should not be present on the page.
        :param parameters: Dictionary of parameters to be used with this test case.
        :param variables: Dictionary of variables stored form 'test_data' sheet
        :param objectRepo: Dictionary of element locators from 'repository' sheet.
        :return: N/A
        """
        elements = self.find_elements(parameters, objectRepo)
        assert elements is None or len(elements) == 0

    def find_element(self, parameters, objectRepo):
        """
        Find an element on a webpage
        :param parameters: Dictionary of parameters to be used with this test case.
        :param objectRepo: Dictionary of element locators from 'repository' sheet.
        :return: Web Element
        """
        locator_type = objectRepo[parameters['Locator']]['Locator_Type']
        locator_value = objectRepo[parameters['Locator']]['Locator_Value']
        wait = WebDriverWait(self.driver, 10)
        try:
            if 'id' in locator_type:
               return wait.until(EC.element_to_be_clickable((By.ID, locator_value)))
            if 'xpath' in locator_type:
                return wait.until(EC.element_to_be_clickable((By.XPATH, locator_value)))
        except Exception as e:
            print('Failed due to - ' +str(e))
            self.tear_down()

    def find_elements(self, parameters, objectRepo):
        """
        Find all the elements by a locator on a webpage.
        :param parameters: Dictionary of parameters to be used with this test case.
        :param objectRepo: Dictionary of element locators from 'repository' sheet.
        :return: List of WebElements.
        """
        locator_type = objectRepo[parameters['Locator']]['Locator_Type']
        locator_value = objectRepo[parameters['Locator']]['Locator_Value']
        if 'id' in locator_type:
            return self.driver.find_elements_by_id(locator_value)
        if 'xpath' in locator_type:
            return self.driver.find_elements_by_xpath(locator_value)






