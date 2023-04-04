from jinja2 import (Environment,
                    Template,
                    meta)

class ToTemplate:
    def __init__(self):
        self.env = Environment()

    def get_template(self, file_path: str):
        """
        todo
        """
        with open(file_path, encoding = 'UTF-8') as file:
            template = Template(file.read())

        return template

    def __get_variables(self, file_path: str) -> list:
        with open(file_path, encoding = 'UTF-8') as file:
            parsed = self.env.parse(file.read())
            variables = meta.find_undeclared_variables(parsed)

        return variables

    def get_variables(self, *args) -> list:
        """
        todo
        """
        variables = []
        for arg in args:
            variables_ = self.__get_variables(arg)
            variables.extend(variables_)

        return variables
