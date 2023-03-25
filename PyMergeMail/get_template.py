import os
from jinja2 import (Environment,
                    FileSystemLoader,
                    select_autoescape,
                    meta)

class template:
    def __init__(self, file_path):
        template_dir = os.path.dirname(file_path)
        self.env = Environment(loader=FileSystemLoader(template_dir),
                          autoescape=select_autoescape())
        self.file_name = os.path.basename(file_path)

    async def get_template(self):
        """
        todo
        """
        return self.env.get_template(self.file_name)

    async def get_vars(self):
        source = self.env.loader.get_source(self.env, self.file_name)
        parced_content = self.env.parse(source)

        return meta.find_undeclared_variables(parced_content)

async def get_variables(*args):
    """
    todo
    """
    variables = []
    for arg in args:
        vars_list = await template(arg).get_vars()
        variables.extend(vars_list)

    return variables
