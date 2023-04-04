from jinja2 import (Environment,
                    Template,
                    meta)

class ToTemplate:
    def __init__(self, file_path: str):
        self.env = Environment()

        with open(f"{file_path}", encoding = 'UTF-8') as file:
            self.template = Template(file.read())

    def get_template(self):
        """
        todo
        """
        return self.template

    def get_variables__(self) -> list:
        parced_content = self.env.parse(self.template)
        variables_ = meta.find_undeclared_variables(parced_content)

        return variables_

def get_variables(*args) -> list:
    """
    todo
    """
    variables = []
    for arg in args:
        variables_ = ToTemplate(arg).get_variables__()
        variables.extend(variables_)

    return variables
