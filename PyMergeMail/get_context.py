from email.utils import make_msgid

async def get_context(data_dict: dict,
                      variables: list = None,
                      cid_fields: list = None):
    """
        obtain required info from excel
    """
    context = {}
    for variable in variables:
        if variable in data_dict:
            context[variable] = data_dict[variable]

    img_path_cid = {}
    for cid_field in cid_fields:
        path = data_dict.get(cid_field)
        img_cid = make_msgid(cid_field)
        img_path_cid[path] = img_cid
        context[cid_field] = img_cid[1:-1]

    return context, img_path_cid
