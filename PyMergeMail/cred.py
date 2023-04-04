import json

def get_cred(cred_file_path, change: bool = False) -> dict:
    """
    print cred on terminal
    """
    with open(cred_file_path, encoding="UTF-8") as cred_file:
        cred: dict = json.load(cred_file)
        print(f"your current credentials: {cred}")

    def smart_input(msg: str, assign_with: str) -> str:
        input_text = input(msg)
        if input_text == "":
            input_text = assign_with

        return input_text

    if change is True:
        input_email = smart_input("Enter email: ", cred['email'])
        input_pass = smart_input("Enter app password: ", cred['pass'])

        with open(cred_file_path, "r+", encoding="UTF-8") as cred_file:
            cred.update({"email":input_email, "pass":input_pass})
            cred_file.seek(0)
            json.dump(cred, cred_file, indent=4)
            cred_file.truncate()
            print(f"New keys updated...\nupdated keys: {cred}")

    return cred
