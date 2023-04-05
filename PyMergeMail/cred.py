import json

class Cred:
    def __init__(self, cred_file_path: str):
        self.cred_file_path = cred_file_path

        with open(cred_file_path, encoding="UTF-8") as cred_file:
            self.cred: dict = json.load(cred_file)
            # print(f"your current credentials: {self.cred}")

    @property
    async def get_cred(self) -> dict:
        """get cred from supplied path
        """
        return self.cred

    @property
    async def update_cred(self) -> dict:
        def smart_input(msg: str, assign_with: str) -> str:
            input_text = input(msg)
            if input_text == "":
                input_text = assign_with

            return input_text

        input_email = smart_input("Enter email: ", self.cred.get('email'))
        input_pass = smart_input("Enter app password: ", self.cred.get('pass'))
        input_alias = smart_input("Enter alias: ", self.cred.get('alias'))

        with open(self.cred_file_path, "r+", encoding="UTF-8") as cred_file:
            self.cred.update({"email":input_email, "pass":input_pass, "alias":input_alias})
            cred_file.seek(0)
            json.dump(self.cred, cred_file, indent=4)
            cred_file.truncate()
            print("New keys updated...")

        return self.cred
