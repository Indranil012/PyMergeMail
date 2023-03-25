import json

def smart_input(msg: str, assign_with: str):
    input_text = input(msg)
    if input_text == "":
        input_text = assign_with

    return input_text

def cred(cred_file_path: str, change: bool = False):
    """
    print cred on terminal
    """
    with open(cred_file_path, encoding="UTF-8") as cred_file:
        cred_dict = json.load(cred_file)
    print(f"your current credentials: {cred_dict}")
    if change is True:
        input_email = smart_input("Enter email: ", cred_dict['email'])
        input_pass = smart_input("Enter app password: ", cred_dict['pass'])

        with open(cred_file_path, "r+", encoding="UTF-8") as cred_file:
            cred_dict.update({"email":input_email, "pass":input_pass})
            cred_file.seek(0)
            json.dump(cred_dict, cred_file, indent=4)
            cred_file.truncate()
        print(f"New keys updated...\nupdated keys: {cred_dict}")

    return cred_dict
