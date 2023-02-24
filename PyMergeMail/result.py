from PyMergeMail.color_print import color_print

async def result(count_successful, count_unsuccessful):
    """
    todo
    """
    total_email = count_successful + count_unsuccessful
    if count_unsuccessful == 0:
        print(color_print(f"{count_successful}/{total_email} Mail successfully sent", 157))
        # 157 for green
    else:
        color_print(f"{count_successful}/{total_email} Mail successfully sent and", 157)
        # 157 for green
        color_print(f"{count_unsuccessful}/{total_email} Mail couldn't be sent", 160)
        # 160 for red
