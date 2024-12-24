import requests
import tqdm

_file = open("2023-07-29/out.txt").readlines()
for link in tqdm.tqdm(_file, total=len(_file)):
    if (status_code := str(requests.head(link).status_code)) != "200":
        print(link)
        print(status_code)
