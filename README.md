# PyMergeMail
A simple module to send bulk personalized html mail with jinja2 template
# Installation
```python
pip install PyMergeMail
```
# Usage
```python
import asyncio
from PyMergeMail import mail

CRED_FILE_PATH = "key.json"
DATA_FILE_PATH = "source_data.xlsx"
SUBJECT_FILE_PATH = "subject.txt"
BODY_FILE_PATH = "body.html"
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
```
