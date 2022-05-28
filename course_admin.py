from __future__ import annotations
from typing import List, Dict, Tuple, Optional
import pathlib
import json
import getpass
import random
import more_itertools
import exchangelib

__version__ = "0.0.3"

Email: str
# interp. An email address as a string
# Examples: 
EM_1 = "cferster@rjc.ca"

Name: str
# interp. A person's full name as registered with their email
# Examples: 
N1 = "Connor Ferster"

UserName: str
# interp. A person's email user name
# Examples:
UN1 = "cferster"


Members: Dict[Email, Name]
# interp: A dict with Email keys and Name values to represent active members of a group
# Examples
M1 = {EM_1: N1}


def send_workbooks_for_review(
    folder_path: List[str],
    account: exchangelib.Account,
    json_path: pathlib.Path,
    email_body: str,
    test_print = True
    ) -> None:
    """
    Returns None. Forwards each email found in the inbox 'folder_name' (as sub folder of inbox),
    to a random recipient chosen from one of the other email senders of the emails found in 'folder_name'.
    Creates a JSON file of email recipient pairings in the CWD. 

    'folder_path': a list of sub-folders to get to the appropriate folder path in the 'account' inbox.
        No need to include the "inbox" folder.
    'account': an ExchangeLib account
    'json_path': a path at which to create the JSON pairings file
    'email_body': the body of the email to send. Must include {person} and {workbook_title} for formatting.
    'test_print': If True, will not send email but will perform all steps and print a test log
    """
    cwd = json_path
    workbook_title = folder_path[-1] # The last folder name
    target_folder = navigate_to_target_folder(folder_path, account)
    if test_print:
        print(f"Target email folder: {target_folder}")
    members = get_email_addresses_from_msgs(target_folder)
    random_pairings = create_random_pairings(members)
    json_dump_pairings(random_pairings, cwd / (workbook_title + ".json"))
    email_subject = f"For Review: {workbook_title}"
    if not test_print:
        for pairing in random_pairings:
            pers_a, pers_b = pairing
            msg_a = next(msg for msg in target_folder.filter(sender=pers_a))
            msg_b = next(msg for msg in target_folder.filter(sender=pers_b))
            forward_email(msg_a, pers_b, email_subject, email_body.format(person=members[pers_a], workbook_title=workbook_title))
            forward_email(msg_b, pers_a, email_subject, email_body.format(person=members[pers_b], workbook_title=workbook_title))
            print(f"Pairing: {pers_a}, {pers_b} -> Sent: {email_subject}")
        return
    else:
        pers_a, pers_b = random_pairings[0]
        print(email_body.format(person=members[pers_a], workbook_title=workbook_title))
        
    

def return_reviewed_notebooks(
    folder_path: List[str],
    account: exchangelib.Account,
    email_body: str,
    json_pairings: pathlib.Path
    ) -> List[Email]:
    """
    Returns a list of "unhappy members", people's email addresses who did not
    receive a reviewed notebook email forward.
    Forwards each email found in the Exchange 'account' 'folder_path' to the recipient who is matched to the 
    email sender as described in 'json_pairings' using the Exchange account, 'account'.

    'folder_path': A list of sub directories to navigate to the correct Exchange Inbox folder. No need
        to include the 'inbox'.

    For example, if an email in 'folder_path' was from anouche@domain.com and "anouche@domain.com" was found
    paired with "riley@domain.com" in the 'json_pairings' list, then the email from anouche@domain.com would
    be forwarded to riley@domain.com.
    """
    email_subject = ("Your workbook, {workbook_title}, reviewed!")

    workbook_title = folder_path[-1]
    target_folder = navigate_to_target_folder(folder_path, account)
    members = get_email_addresses_from_msgs(target_folder)
    pairings = json_load_pairings(json_pairings)
    email_subject = f"{workbook_title} Review"
    members_receiving_emails = []
    no_matches = []
    for member in members.keys():
        pers_b = find_pair_match(member, pairings)
        if not pers_b: 
            no_matches.append(member)
            continue
        returned_workbook = next(msg for msg in target_folder.filter(sender=member))
        forward_email(
            returned_workbook, 
            pers_b, 
            email_subject.format(workbook_title), 
            email_body.format(name=members[member], workbook_title=workbook_title)
        )
        members_receiving_emails.append(pers_b)
    all_members = set([person for pairing in pairings for person in pairing])
    happy_members = set(members_receiving_emails)
    unhappy_members = all_members - happy_members
    print("Unhappy members: ", unhappy_members)
    return unhappy_members


def email_unhappy_members(
    unhappy_members: List[Email], 
    workbook_title: str, 
    account: exchangelib.Account, 
    json_pairings: pathlib.Path,
    ) -> None:
    """
    Returns None. Emails the people in 'unhappy_members' who submitted a workbook for review but did not receive a returned
    reviewed notebook.
    """
    if not unhappy_members:
        print("No unhappy members :)")
        return
    alt_email_subject = ("Your workbook, {workbook_title}, was not reviewed :(") 
    alt_email_body = (
        "Your reviewer was not able to return your reviewed notebook (probably busy, yeah?)."
        " Contact them directly at {email} if you would like to see their review."
    )
    pairings = json_load_pairings(json_pairings)
    for member in unhappy_members:
        reviewers_email = find_pair_match(member, pairings)
        msg = exchangelib.Message(
            account=account,
            subject=alt_email_subject.format(workbook_title=workbook_title),
            body=alt_email_body.format(email=reviewers_email),
            to_recipients=[member]
        )
        msg.send()
        # Have to create a new email from scratch
        

def get_email_addresses_from_msgs(msg_folder: exchangelib.folders.known_folder.Messages) -> Members:
    """
    Returns a Members dict representing the email addresses and names of the senders for all the messages
    contained in 'msg_folder'
    """
    mems = {}
    for msg in msg_folder.all():
        mems.update({msg.sender.email_address: msg.sender.name})
    return mems
            

def navigate_to_target_folder(folder_path: List[str], account: exchangelib.Account) -> exchangelib.folders.known_folders.Messages:
    """
    Returns a reference to the target folder at the end of 'folder_path' in the 'account' inbox.

    'folder_path': a list of str representing the heirarchy of folders and sub folders. e.g. 
    ["Python_Course", "Workbook 1"] would navigate to account.inbox / "Python Course" / "Workbook 1"
    """
    target_folder = account.inbox
    for folder in folder_path:
        target_folder = target_folder / folder
    return target_folder



def create_random_pairings(submissions: Members, filler: Optional[Email] = None) -> List[Tuple[Email, Email]]:
    """
    Returns a list of tuple pairs representing random pairings of the member
    email addresses in 'submissions'. If there is an odd number of pairings then
    the last member is matched with the value in 'filler'.
    """
    emails = list(submissions.keys())
    random.shuffle(emails)
    random_pairs = []
    for pair in more_itertools.grouper(emails, 2, filler):
        random_pairs.append(pair)
    return random_pairs


def find_pair_match(email_a: Email, pairings: List[List[Email, Email]]) -> Optional[Email]:
    """
    Returns the match, email_b, for corresponding 'email_a' in 'pairings'.
    Returns None if no match found.
    """
    for pairing in pairings:
        pers_a, pers_b = pairing
        if pers_a == email_a:
            return pers_b
        elif pers_b == email_a:
            return pers_a
    else:
        return


def json_dump_pairings(pairings: List[tuple], file_path: pathlib.Path) -> None:
    """
    Returns None. Saves the list of pairings to a JSON file at 'file_path'.
    """
    with open(file_path, 'w') as file:
        json.dump(pairings, file)
    return


def json_load_pairings(file_path: pathlib.Path) -> List[List[Email, Email]]:
    """
    Returns a list of email pairings contained in the json file, 'file_path'.
    """
    with open(file_path, "r") as file:
        pairings = json.load(file)
    return pairings


def get_email_user(email: Email) -> str:
    """
    Returns the username portion of an email address.
    e.g. email="cferster@rjc.ca" returns "cferster"
    """
    return email.split("@")[0].lower()


def connect_to_rjc_exchange(email: Email, username: UserName) -> exchangelib.Account:
    """
    Get Exchange account connection with RJC email server
    """
    server = "mail.rjc.ca"
    domain = "RJC"
    return connect_to_exchange(server, domain, email, username)


def connect_to_exchange(server: str, domain: str, email: str, username: str) -> exchangelib.Account:
    """
    Get Exchange account connection with server
    """
    credentials = exchangelib.Credentials(username= f'{domain}\\{username}', 
                                          password=getpass.getpass("Exchange pass:"))
    # config = exchangelib.Configuration(server=server, credentials=credentials)
    config = exchangelib.Configuration(server='outlook.office365.com', credentials=credentials)

    return exchangelib.Account(primary_smtp_address=email, autodiscover=False, 
                   config = config, access_type=exchangelib.DELEGATE)


def forward_email(msg: exchangelib.Message, fwd_to: Email, subject: str, body: Optional[str] = None) -> None:
    """
    Returns None. Forwards the email message, 'msg', to the email address, 'fwd_to' with a
    subject line of 'subject' and with body text as 'body'.
    """
    msg.forward(subject=subject, body=body, to_recipients=[fwd_to])
    return








