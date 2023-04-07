"""
sample usecase
"""
import asyncio
from PyMergeMail import mail

CRED_FILE_PATH = "sample_data/key.json"
DATA_FILE_PATH = "sample_data/example_data.xlsx"
SUBJECT_FILE_PATH = "sample_data/subject.txt"
BODY_FILE_PATH = "sample_data/body.html"
CID_FIELDS = ["img_path", "sig_path"]
ATTACH_FIELD = "attachment"

asyncio.run(mail(CRED_FILE_PATH,
                 DATA_FILE_PATH,
                 SUBJECT_FILE_PATH,
                 BODY_FILE_PATH,
                 CID_FIELDS,     # optional
                 ATTACH_FIELD    # optional
            ), debug=True
        )
