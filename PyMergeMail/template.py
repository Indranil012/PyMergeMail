from jinja2 import (Environment,
                    Template,
                    meta)

class ToTemplate:
    env = Environment(enable_async = True)

    @classmethod
    async def get_template(cls, file_path: str):
        """
        todo
        """
        with open(file_path, encoding = 'UTF-8') as file:
            template = Template(file.read(), autoescape = True, enable_async = True)

        return template

    @classmethod
    async def __get_variables(cls, file_path: str) -> list:
        with open(file_path, encoding = 'UTF-8') as file:
            parsed = cls.env.parse(file.read())
            variables = meta.find_undeclared_variables(parsed)

        return variables

    @classmethod
    async def get_variables(cls, *args) -> list:
        """
        todo
        """
        variables = []
        for arg in args:
            variables_ = await cls.__get_variables(arg)
            variables.extend(variables_)

        return variables
