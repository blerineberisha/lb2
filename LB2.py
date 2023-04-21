import json
import os
import platform
import re
import shutil
import sqlite3
import pywhatkit as pwk
import pywinctl
import xml.dom.minidom as xml
from pathlib import Path
import glob

# define global variables that are used consistently throughout the program for the mac actions
userVar = os.environ.get("USER")
original_mac = (
        "/Users/"
        + str(userVar)
        + "/Library/Application Support/AddressBook/Sources/60145976-57A8-4AAF-A0BE-840CA2C577E1/AddressBook-v22.abcddb"
)
dest_mac = "/Users/" + str(userVar) + "/Desktop/AddressBook.abcddb"
menu = "1. Instantly send message to chosen contact\n" \
       "2. Send message to contact at later time today\n"

user = os.getlogin()
win_path = Path("C:/Users/" + str(user) + "/Contacts/").absolute()


# get the operating system, the device is using
def get_os():
    return platform.system()


def os_actions():
    # patterns that match the operating system that this script is running on
    macPattern = re.compile("Darwin.*", flags=re.IGNORECASE)
    linPattern = re.compile("Linux.*", flags=re.IGNORECASE)
    winPattern = re.compile("win.*", flags=re.IGNORECASE)

    # start different functions according to operating system
    opSys = get_os()
    if macPattern.match(opSys):
        return 0
    elif linPattern.match(opSys):
        return 0
    elif winPattern.match(opSys):
        return 0
    else:
        print("Operating system not recognized. No further action possible.")


def mac_actions():
    mac_prep_db()
    print(menu)
    choice = input("How would you like to send a message? ")

    if choice == "1":
        mac_send_message_now()
    elif choice == "2":
        mac_send_message_later()
    else:
        print("Value does not match a choice from the menu. Try again.")


def win_actions():
    return 0


def lin_actions():
    return 0


# This function finds the address book on the macOS device, establishes a connection to the database and returns the
# json output
def get_address_book(address_book_location):
    conn = sqlite3.connect(address_book_location)
    cursor = conn.cursor()

    cursor.execute(
        "SELECT DISTINCT ZABCDRECORD.ZFIRSTNAME, ZABCDRECORD.ZLASTNAME, ZABCDPHONENUMBER.ZFULLNUMBER FROM ZABCDRECORD LEFT JOIN ZABCDPHONENUMBER ON ZABCDRECORD.Z_PK = ZABCDPHONENUMBER.ZOWNER ORDER BY ZABCDRECORD.ZLASTNAME, ZABCDRECORD.ZFIRSTNAME, ZABCDPHONENUMBER.ZORDERINGINDEX"
    )
    result_set = cursor.fetchall()

    json_output = json.dumps(
        [
            {"FIRST NAME": t[0], "LAST NAME": t[1], "FULL NUMBER": t[2]}
            for t in result_set
        ]
    )

    conn.close()
    return json_output


# due to access issues, the file needs to be copied to a place where the rights to read, write and execute are
# given. here it's the desktop.
def mac_copy_address_book():
    open("/Users/" + userVar + "/Desktop/AddressBook.abcddb", "w")
    # copy database to new destination
    shutil.copy(original_mac, dest_mac)
    get_address_book(address_book_location=dest_mac)
    return dest_mac


# finds a contact in the database based on the first name saved in the device
def find_contact():
    contact_first = input("What is the contacts first name? ")
    conn = sqlite3.connect(dest_mac)
    cur = conn.cursor()
    # query that gets the phone number with the number code based on the entered name. this query is case-insensitive
    cur.execute(
        "SELECT ZABCDPHONENUMBER.ZFULLNUMBER FROM ZABCDRECORD "
        "LEFT JOIN ZABCDPHONENUMBER ON ZABCDRECORD.Z_PK = ZABCDPHONENUMBER.ZOWNER "
        "WHERE ZABCDRECORD.ZFIRSTNAME=? "
        "COLLATE NOCASE "
        "ORDER BY ZABCDRECORD.ZLASTNAME, ZABCDRECORD.ZFIRSTNAME, ZABCDPHONENUMBER.ZORDERINGINDEX",
        [contact_first])
    rows = cur.fetchall()
    res = 0
    if len(rows) == 1:
        res = rows[0]

    # if there is more than one result, the user is asked to enter the contact's last name, to determine which user
    # the message is supposed to go to
    if len(rows) > 1:
        contact_last = input("What is the contact's last name? ")
        cur.execute(
            "SELECT ZABCDPHONENUMBER.ZFULLNUMBER FROM ZABCDRECORD "
            "LEFT JOIN ZABCDPHONENUMBER ON ZABCDRECORD.Z_PK = ZABCDPHONENUMBER.ZOWNER "
            "WHERE ZABCDRECORD.ZFIRSTNAME=? AND ZABCDRECORD.ZLASTNAME=? "
            "COLLATE NOCASE "
            "ORDER BY ZABCDRECORD.ZLASTNAME, ZABCDRECORD.ZFIRSTNAME, ZABCDPHONENUMBER.ZORDERINGINDEX",
            [contact_first, contact_last])
        rows = cur.fetchall()
        for row in rows:
            res = row
    return res


def send_message():
    number = str(find_contact())
    message = input("Enter the message you'd like to send:")
    time = input("Enter the desired time to send the message:")
    time_arr = time.split(":")
    hour = int(time_arr[0])
    mins = int(time_arr[1])
    try:
        # sending message with chosen contact and entered message at desired time, including waiting time
        ''' 15s of waiting time (4th argument) is required, since whatsapp web takes a few
        seconds to load. if the local internet is not sufficiently fast, change the waiting
        time to 30s or more.'''
        pwk.sendwhatmsg(number, message, hour, mins, 15)
        print("Message Sent!")  # Prints success message in console
    except Exception as e:
        print("Error in sending the message: ", {e})


def mac_prep_db():
    get_address_book(original_mac)
    mac_copy_address_book()


def mac_send_message_later():
    send_message()


def mac_send_message_now():
    number = str(find_contact())
    message = input("Enter the message you'd like to send:")
    try:
        # sending message with chosen contact and entered message at desired time, including waiting time
        ''' 15s of waiting time (4th argument) is required, since whatsapp web takes a few
        seconds to load. if the local internet is not sufficiently fast, change the waiting
        time to 30s or more.'''
        pwk.sendwhatmsg_instantly(number, message, 15)
        pywinctl.getWindowsWithTitle("WhatsApp")
        print("Message Sent!")  # Prints success message in console
    except Exception as e:
        print("Error in sending the message: ", {e})


# My path for the contacts is: "C:\Users\bleri\Contacts", so the general
# path should be "C:\Users\[username]\Contacts\" and there you should find your
# .contact files.
def win_get_contact_files():
    os.chdir(win_path)
    return glob.glob("*.contact")


def win_get_nr_from_contact():
    all_contacts = win_get_contact_files()
    # recipient = input("Which contact would you like to send a message to?")
    for c in all_contacts:
        doc = xml.parse(c)
        print(doc)


win_get_nr_from_contact()
