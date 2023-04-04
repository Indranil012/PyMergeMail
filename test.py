"""
sample usecase
"""
import asyncio
from PyMergeMail import mail

CRED_FILE_PATH = "source_data/key.json"
DATA_FILE_PATH = "source_data/source_data.xlsx"
SUBJECT_FILE_PATH = "source_data/subject.txt"
BODY_FILE_PATH = "source_data/test.html"
CID_FIELDS = ["img_path", "sig_path"]
ATTACH_FIELD = "attachment"

asyncio.run(mail(CRED_FILE_PATH,
                 DATA_FILE_PATH,
                 SUBJECT_FILE_PATH,
                 BODY_FILE_PATH,
                 CID_FIELDS,
                 ATTACH_FIELD), debug=True)
