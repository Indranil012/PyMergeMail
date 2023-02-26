"""
sample usecase
"""
import asyncio
from PyMergeMail import send

if __name__ == "__main__":
    CRED_FILE_PATH = "source_data/key.json"
    DATA_FILE_PATH = "source_data/source_data.xlsx"
    SUBJECT_FILE_PATH = "source_data/subject.txt"
    BODY_FILE_PATH = "source_data/test.html"
    CID_FIELDS = ["img_path", "sig_path"]
    ATTACH_FIELD = "attachment"

    asyncio.run(send(CRED_FILE_PATH,
                          DATA_FILE_PATH,
                          SUBJECT_FILE_PATH,
                          BODY_FILE_PATH,
                          CID_FIELDS,
                          ATTACH_FIELD), debug=True)
