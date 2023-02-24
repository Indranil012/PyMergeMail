from PyMergeMail.get_template import get_template

async def get_var(subject_file_path,
                  body_file_path):
    """
    todo
    """
    body_var = await get_template(body_file_path, var_only=True)
    subject_var = await get_template(subject_file_path, var_only=True)

    return body_var | subject_var
