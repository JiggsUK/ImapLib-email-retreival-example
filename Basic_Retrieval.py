#!/usr/bin/env python
"""
Created by Jiggs

Basic python script to connect to an IMAP mailbox, retrieve a specific line of text from an email body and write it to a
csv file.
Includes various functions for cleaning up the returned email body if necessary.
Intended as a base code to assist with building an email checking program and provide an example of what you can do with
your emails.
"""

import imaplib
import re
from _datetime import datetime
import itertools
import threading
import time
import sys

# connecting to your email
mail = imaplib.IMAP4_SSL('imap.google.com')  # imap server address such as google
mail.login('youremail@gmail.com', 'password')  # strongly advise against including password in code for security reasons

list_response_pattern = re.compile(r'\((?P<flags>.*?)\) "(?P<delimiter>.*)" (?P<name>.*)')


def parse_list_response(line):
    """
    Function to clean up mailbox list to create a user-friendly, readable list.
    """
    match = list_response_pattern.match(line.decode('utf-8'))
    flags, delimiter, mailbox_name = match.groups()
    mailbox_name = mailbox_name.strip('"')
    return flags, delimiter, mailbox_name


def list_mailboxes(mailbox):
    """
    Works in tandem with the 'parse_list_response' function to list the available mailboxes using imaplib's list method.
    More of a debug feature if mailbox names are changed in future - mailbox names should be one word and ARE case
    sensitive when calling the select method.
    """
    response_code_list, data_list = mailbox.list()
    if response_code_list != "OK":
        print("There was an error. Please try again.")
    else:
        print('Your mailboxes: ')
        for line in data_list:
            user_mailbox_as_a_list = []
            flags, delimiter, mailbox_name = parse_list_response(line)
            user_mailbox_as_a_list.append(mailbox_name)  # adds mailbox names to the empty list created above
            print(" ".join(user_mailbox_as_a_list))


def is_empty(any_structure):
    """
    Checks to see if the list of unseen email IDs (list_check) is empty.
    This function is used to enter the while loop and iterate over the emails.
    If it's empty there are no new emails and therefore the script does not need to run and the initial if statement
    will run instead.
    """
    if any_structure:
        # structure is not empty
        return False
    else:
        # structure is empty
        return True


def search_string(data_fetched):
    """
    Function to search the returned email string for the word 'Hi' and return the words after it.
    Chose to collect 90 expressions after the keyword to ensure variance in the length of ship names and locations are
    accounted for.
    It returns the matched string as a list 'returned_string_as_list'
    """
    search = re.search(r'((Hi)\W+((?:\w+\W+){,90}))', data_fetched)
    if search:
        new_list = [x.strip().split() for x in search.groups()]
        returned_string_as_list = new_list[2]
        return returned_string_as_list


def write_to_file(item):
    """
    Function to write the wanted information to a csv file. It adds the current month to the filename which means you
    will get a new file on the 1st of every month.
    It will append any new sentences to the file for the current month, each time the script runs.
    """
    current_date = datetime.now()
    current_month = current_date.strftime("%b")
    f = open("Your Filename %s.csv" % current_month, "a+")
    f.write(item + "\n")  # write each item to the file with a line break at the end - will write each item on a new line
    f.close()


list_mailboxes(mail)  # view the available mailboxes to select one

chosen_mailbox = mail.select("INBOX")  # select the mailbox to go to
response_code, data = mail.search(None, 'UNSEEN')  # search for unread emails
list_check = data[0].split()  # drop email IDs into a list - used to break the while loop below
if response_code != "OK":
    print("Problem reaching mailbox: Inbox")  # will tell you if it can't reach mailbox

if is_empty(list_check) is True:
    print("\r No unread messages in mailbox: Inbox")  # if there are no unread emails it will not enter the loop

current_date_for_header = datetime.today()
if current_date_for_header.day == 1:  # filename changes on the first of the month - will only print header row then
    header_row = ['Col 1', 'Col 2', 'Col 3']
    header_row_to_print = ", ".join(header_row)
    write_to_file(str(header_row_to_print))  # prints header row to file before the emails

while is_empty(list_check) is False:  # checks list_check for unread emails, enters the loop if there is
    for num in list_check:  # for each email ID in list_check run the following code on that email
        response_code_fetch, data_fetch = mail.fetch(num, '(BODY[TEXT])')  # fetch the text from the body of the email
        if response_code_fetch != "OK":
            print("Unable to find requested messages")

        searched_email = search_string(str(data_fetch))
        final_result = ', '.join(searched_email)  # join list using comma delimiter
        write_to_file(str(final_result))  # write list to the csv file
        typ_stored, data_stored = mail.store(num, '+FLAGS', '\\Seen')  # marks as read - to prevent double counts later
        list_check.remove(num)  # removes the searched email ID from the list - prevents an infinite loop
