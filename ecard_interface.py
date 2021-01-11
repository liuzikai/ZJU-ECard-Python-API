"""
This is a module to acquire transactions data from ecardhall.zju.edu.cn.
"""

import requests
import json
import time
import skimage.io
import skimage.transform
import math
from Crypto.PublicKey import RSA
from Crypto.Cipher import PKCS1_v1_5
import binascii


class ECardInterface:

    def __init__(self):
        self.session = requests.session()
        self.ecard_account = None

    def get_checkcode(self) -> bytes:
        """
        Get (refresh) checkcode image.
        :return: Checkcode image content in bytes
        """

        # Access main page
        ret = self.session.get("http://ecardhall.zju.edu.cn:808").status_code
        assert ret == 200, "Failed to access main page"

        # Generate checkcode time stamp
        checkcode_flag = str(int(time.time() * 1000))

        # Get checkcode
        checkcode_ret = self.session.get("http://ecardhall.zju.edu.cn:808/Login/GetValidateCode?time=" + checkcode_flag)
        assert checkcode_ret.status_code == 200, "Failed to get checkcode"

        return checkcode_ret.content

    @staticmethod
    def display_checkcode(img_data: bytes) -> None:
        """
        Read image from file and display the image to the terminal using colored char.
        :param img_data: the image data in bytes
        :return: None
        """
        # Print checkcode to CUI
        checkcode_img = skimage.io.imread(img_data, as_gray=True, plugin='imageio')
        # checkcode_img = skimage.transform.rescale(checkcode_img, 0.5, mode="edge")
        rows, cols = checkcode_img.shape
        for i in range(rows):
            for j in range(cols):
                if checkcode_img[i, j] <= 0.7:
                    print("\033[7m%s\033[0m" % "  ", sep="", end="")  # black
                elif checkcode_img[i, j] <= 0.8:
                    print("\033[0;47m%s\033[0m" % "  ", sep="", end="")  # grey
                else:
                    print("\033[0;40m%s\033[0m" % "  ", sep="", end="")  # white
            print("")

    def login(self, ecard_account: str, ecard_secret: str, checkcode: str) -> bool:
        """
        Login
        :param ecard_account:
        :param ecard_secret:
        :param checkcode: checkcode number entered from user
        :return: True if login succeeded. False if failed because of checkcode error. Error is raised if login failed
                 because of other error.
        """

        shared_headers = {
            "Accept": "application/json, text/javascript, */*; q=0.01",
            "Content-Type": "application/x-www-form-urlencoded; charset=UTF-8",
            "Referer": "http://ecardhall.zju.edu.cn:808/",
            "Origin": "http://ecardhall.zju.edu.cn:808",
            "Host": "ecardhall.zju.edu.cn:808",
            "Accept-Language": "en-us",
            "Accept-Encoding": "gzip, deflate",
            "X-Requested-With": "XMLHttpRequest",
            "Connection": "keep-alive"
        }

        rsa_key_data = {
            "MIME Type": "application/x-www-form-urlencoded; charset=UTF-8",
            "json": "true"
        }

        ret = self.session.post("http://ecardhall.zju.edu.cn:808/Common/GetRsaKey",
                                headers=shared_headers,
                                data=rsa_key_data)

        assert ret.status_code == 200, "Failed to access page of getting RSA key"

        rsa_key_json = json.loads(ret.text)

        assert rsa_key_json["IsSucceed"], "Failed to get RSA key"

        pub_key = RSA.construct(
            (int(rsa_key_json["Obj"].split(",")[1], 16), int(rsa_key_json["Obj"].split(",")[0], 16)))
        pub_key = PKCS1_v1_5.new(pub_key)

        login_data = {
            "MIME Type": "application/x-www-form-urlencoded; charset=UTF-8",
            "sno": ecard_account,
            "pwd": binascii.hexlify(pub_key.encrypt(ecard_secret.encode())).decode("utf-8"),
            "ValiCode": checkcode,
            "remember": "1",
            "uclass": "1a",
            "zqcode": "",
            "key": rsa_key_json["Msg"],
            "json": "true"
        }
        ret = self.session.post("http://ecardhall.zju.edu.cn:808/Login/LoginBySnoQuery",
                                headers=shared_headers,
                                data=login_data)

        assert ret.status_code == 200, "Failed to access login page"

        login_json = json.loads(ret.text)

        if login_json["IsSucceed"]:
            self.ecard_account = ecard_account
            return True

        if login_json["Msg"] == "验证码错误":
            return False

        print(ret.text)
        raise Exception("Unexpected error during login")

    def _fetch_data_raw(self, ecard_account: str, begin_date: str, end_date: str, page: int) -> str:
        """
        Fetch raw data of transaction. See test data for its format.
        :param begin_date: string in "YYYY-MM-DD"
        :param end_date: string in "YYYY-MM-DD"
        :param page: page index
        :return: string of raw data of transaction.
        """

        fetch_headers = {
            "Accept": "application/json, text/javascript, */*; q=0.01",
            "Content-Type": "application/x-www-form-urlencoded; charset=UTF-8",
            "Referer": "http://ecardhall.zju.edu.cn:808/Page/Page",
            "Origin": "http://ecardhall.zju.edu.cn:808",
            "Host": "ecardhall.zju.edu.cn:808",
            "Accept-Language": "en-us",
            "Accept-Encoding": "gzip, deflate",
            "X-Requested-With": "XMLHttpRequest",
            "Connection": "keep-alive"
        }

        fetch_data = {
            "sdate": begin_date,
            "edate": end_date,
            "account": ecard_account,
            "page": str(page),
            "rows": "100"
        }

        ret = self.session.post("http://ecardhall.zju.edu.cn:808/Report/GetPersonTrjn",
                                headers=fetch_headers,
                                data=fetch_data)
        assert ret.status_code == 200, "Failed to fetch raw data"
        return str(ret.content, "utf-8")

    def acquire_data(self, ecard_account: str, begin_date: str, end_date: str) -> list:
        """
        Main function to acquire and process data
        :param ecard_account: ZJU ecard secret in string
        :param begin_date: string in "YYYY-MM-DD"
        :param end_date: string in "YYYY-MM-DD"
        :return: list of transaction in the format of [Time, Type, Subtype, Value, Remainder] (all string type)
        """

        assert self.ecard_account is not None, "Haven't logged in yet"

        page_index = 1
        page_count = 1
        record_entries = []

        while page_index <= page_count:
            page_json = json.loads(self._fetch_data_raw(ecard_account, begin_date, end_date, 1))
            if page_index == 1:
                # Set the real page count based on the data on page 1
                page_count = math.ceil(int(page_json["total"]) / 100)
            for row in page_json["rows"]:
                record_entries.append(
                    [
                        row["OCCTIME"].strip(),
                        row["MERCNAME"].strip(),
                        row["TRANNAME"].strip(),
                        str(row["TRANAMT"]),
                        str(row["CARDBAL"])
                    ])
            page_index += 1

        return record_entries
