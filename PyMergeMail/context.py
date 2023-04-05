from email.utils import make_msgid

class Context:
    def __init__(self,
                 recv_data: dict,
                 variables: list,
                 cid_fields: list):

        self.context = {}
        self.img_path_cid = {}

        for variable in variables:
            if variable in recv_data:
                self.context[variable] = recv_data.get(variable)

        if cid_fields is not None:
            for cid_field in cid_fields:
                path = recv_data.get(cid_field)
                img_cid = make_msgid(cid_field)
                self.img_path_cid[path] = img_cid
                self.context[cid_field] = img_cid[1:-1]

    @property
    async def get_context(self) -> dict:
        """
            obtain required info from excel
        """
        return self.context

    @property
    async def get_img_cid(self) -> dict:
        return self.img_path_cid
