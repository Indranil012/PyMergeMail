from email.utils import make_msgid

async def get_context(row,
                      variables,
                      cid_fields):
    """
        obtain required info from excel
    """
    context = {}
    for variable in variables:
        if variable in row:
            context[variable] = row[variable]

    img_path_cid = {}
    for cid_field in cid_fields:
        path = row.get(cid_field)
        img_cid = make_msgid(cid_field)
        img_path_cid[path] = img_cid
        context[cid_field] = img_cid[1:-1]

    return context, img_path_cid
