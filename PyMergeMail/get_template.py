import os
from jinja2 import (Environment,
                    FileSystemLoader,
                    select_autoescape,
                    meta)

async def get_template(file_path, var_only=False):
    """
    todo
    """
    template_dir = os.path.dirname(file_path)
    env = Environment(
        loader=FileSystemLoader(
            template_dir),
                      autoescape=select_autoescape())
    file_name = os.path.basename(file_path)

    if var_only is True:
        source = env.loader.get_source(env, file_name)
        parced_content = env.parse(source)
        return meta.find_undeclared_variables(parced_content)

    return env.get_template(file_name)
