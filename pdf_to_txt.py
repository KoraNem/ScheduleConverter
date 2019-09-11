import cloudconvert
from os import path

api_key = 'asWrUsZVGa7xI0GRcgMOWpBxVMsRWHBqdY0VpwdyXwoLYJqIPK7eXuPa4UmZS5kQ'


def pdf_txt(file_path):
    if path.exists(file_path):
        print("Ok!", file_path)

    api = cloudconvert.Api(api_key)

    process = api.convert({
        "inputformat": "pdf",
        "outputformat": "txt",
        "input": "upload",
        "file": open(file_path, 'rb')
    })
    process.wait()
    process.download()
