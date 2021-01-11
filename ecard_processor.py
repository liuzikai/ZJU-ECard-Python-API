import datetime
import configparser
import os.path

from ecard_interface import ECardInterface
import excel_generate

from termcolor import colored, cprint

config = configparser.ConfigParser()
config_filename = os.path.join(os.path.dirname(__file__), "config.ini")
config.read(config_filename)


def get_end_date() -> (str, bool):
    ret = (datetime.date.today() - datetime.timedelta(days=1)).strftime("%Y-%m-%d")
    manual_set = False
    print("Today is %s. " % datetime.date.today().strftime("%Y-%m-%d"), end="")
    ipt = input(colored("Press enter to use %s as end date, or input manually: " % ret, "yellow"))
    if ipt != '':
        ret = ipt
        manual_set = True
    return ret, manual_set


def get_begin_date() -> (str, bool):
    manual_set = False
    try:
        last_date = config.get("History", "LastDate")
        ret = (datetime.datetime.strptime(last_date, "%Y-%m-%d") + datetime.timedelta(days=1)).strftime(
            "%Y-%m-%d")
        print("Last time you convey the record till %s." % last_date)
        ipt = input(colored("Press enter to use %s as start date, or input manually: " % ret, "yellow"))
        if ipt != '':
            ret = ipt
            manual_set = True
    except configparser.NoSectionError or config.NoOptionError:
        ret = input("Can't find usage history. Please type in the begin date (YYYY-MM-DD): ")
        manual_set = True
    return ret, manual_set


def get_workbook_filename() -> str:
    try:
        workbook_path = config["Settings"]["Workbook_Path"]
    except configparser.NoSectionError or config.NoOptionError:
        workbook_path = input("Please set output path: ")
        if "Settings" not in config:
            config.add_section("Settings")
        config.set("Settings", "Workbook_Path", workbook_path)
        config.write(open('config.ini', 'w'))

    return workbook_path + "/" + datetime.datetime.now().strftime("%Y-%m-%d-%H-%M-%S") + ".xls"


def get_account() -> str:
    try:
        account = config["Settings"]["ecard_account"]
    except configparser.NoSectionError or config.NoOptionError:
        account = input("Please set ecard account (6 digits): ")
        if "Settings" not in config:
            config.add_section("Settings")
        config.set("Settings", "ecard_account", account)
        config.write(open('config.ini', 'w'))
    return account


def get_secret() -> str:
    try:
        secret = config["Settings"]["ecard_secret"]
    except configparser.NoSectionError or config.NoOptionError:
        secret = input("Please set ecard secret (6 digits): ")
        if "Settings" not in config:
            config.add_section("Settings")
        config.set("Settings", "ecard_secret", secret)
        config.write(open('config.ini', 'w'))
    return secret


def save_config(end_date: str) -> None:
    if "History" not in config:
        config.add_section("History")
    config.set("History", "LastDate", end_date)
    config.write(open(config_filename, "w"))


def run() -> None:
    ecard = ECardInterface()

    # Load data
    ecard_account = get_account()
    ecard_secret = get_secret()
    begin_date, manual_set_begin = get_begin_date()
    end_date, manual_set_end = get_end_date()

    # Input checkcode
    while True:
        ecard.display_checkcode(ecard.get_checkcode())
        checkcode = input(colored("Please input the check code: ", "yellow"))
        if len(checkcode) == 5 and checkcode.isdecimal():
            if ecard.login(ecard_account, ecard_secret, checkcode):
                cprint("Login succeeded", "green")
                break
            else:
                cprint("Error checkcode. CheckCode refreshed.", "yellow")
        else:
            cprint("Invalid Value. CheckCode refreshed.", "yellow")

    # Get records
    records = ecard.acquire_data(ecard_account, begin_date, end_date)

    # Print record
    records_str = ""
    for record in records:
        records_str += " ".join(record) + '\n'
    print(records_str)

    workbook_filename = get_workbook_filename()
    excel_generate.generate_excel(records, workbook_filename)
    cprint("Excel generated to %s" % workbook_filename, "green")

    if not (manual_set_begin or manual_set_end) or input("Save end date? [y]") == "y":
        save_config(end_date)
        cprint("End date saved")

    cprint("https://www.sui.com/data/standard_data_import.do", "blue")


if __name__ == "__main__":
    run()
