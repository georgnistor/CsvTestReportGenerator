class TestCaseGit:
    def __init__(self, tca):
        self._testCaseName = tca


class ModuleGit:
    def __init__(self, name):
        self._name = name

        """"List of Test Cases"""
        self._listTestCases = []

    def append_test_case(self, tca):
        self._listTestCases.append(tca)


class ReportDataGit:
    def __init__(self):
        """List of Modules"""
        self._listModules = []

    def append_module(self, module):
        self._listModules.append(module)
