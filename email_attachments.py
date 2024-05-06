import pathlib
import random
import smtplib
from email.mime.multipart import MIMEMultipart 
from email.mime.text import MIMEText 
from email.mime.base import MIMEBase 
from email import encoders 

import numpy as np
import os
import pandas as pd
from tqdm.auto import tqdm


# read credentials from text file
def read_credentials(filepath='address_password.txt'):
    with open(filepath, 'r') as file:
        # read each line
        for line in file:
            # split line into key/value pair
            key, value = line.strip().split(': ')
            # check if key matches 'address' and store
            if key == 'address':
                sender_address = value
            # check if key matches 'password' and store
            elif key == 'password':
                sender_password = value
    
    return sender_address, sender_password


# get 3 unsent attachments
def sample_attachments(candidates_all, candidates_sent=None, num_samples=3):
    if candidates_sent == None:
        potentials = set(candidates_all)
    else:
        potentials = set(candidates_all) - set(candidates_sent)
        if len(potentials) < num_samples:
            return 'sparse'

    selected = random.sample(list(potentials), num_samples)
    return selected


def main():
    # load record of names and emails
    directory = pd.read_excel('send_files_to.xlsx', header=None, names=['name', 'email', 'notes', 'notes2'])
    names = directory['name'].tolist()
    emails = directory['email'].tolist()

    # read (or initialize) record of loops sent to each email
    if not os.path.exists('sent_files.csv'):
        df_sent = pd.DataFrame({
            'name': names,
            'email': emails,
            'sent_files': [''] * len(names)  # initialize with empty str
        })

        df_sent.to_csv('sent_files.csv', header=True, index=False)

    df_sent = pd.read_csv('sent_files.csv')
    dict_sent = dict(zip(df_sent['email'], df_sent['sent_files']))

    # send emails
    # read credentials from text file
    sender_address, sender_password = read_credentials()

    for receiver_address in tqdm(emails):
        receiver_name = names[emails.index(receiver_address)]

        # instance of MIMEMultipart 
        msg = MIMEMultipart() 
        msg['From'] = sender_address
        msg['To'] = receiver_address
        msg['Subject'] = 'New Files'
        body = ''
        msg.attach(MIMEText(body, 'plain'))

        # get all available attachment files
        files_dir = pathlib.Path(r'attachments')
        filepaths = list(files_dir.glob('*'))
        filenames = [file_path.name for file_path in filepaths if file_path.is_file()]

        if df_sent[df_sent['email'] == receiver_address].index.size == 0:
            new_row = pd.DataFrame(
                {'name': receiver_name, 'email': receiver_address, 'sent_files': np.nan},
                index=[df_sent.index.max() + 1])
            df_sent = pd.concat([df_sent, new_row], ignore_index=True)

        index = df_sent[df_sent['email'] == receiver_address].index[0]  # df row of recipient
        already_sent = df_sent.at[index, 'sent_files']

        # randomly pick 3, write new filenames to df record, update dictionary
        if pd.isna(already_sent):
            selected = sample_attachments(filenames, num_samples=3)
            # if sent loops is NaN, create new list
            updated_sent = ', '.join(selected)
            dict_sent[receiver_address] = selected
        else:
            selected = sample_attachments(filenames, dict_sent[receiver_address].split(', '), num_samples=3)
            if selected == 'sparse':
                print(f'error: not enough unsent files to {receiver_address}')
                continue
            # otherwise append new files to existing string
            updated_sent = ', '.join([already_sent] + selected)
            dict_sent[receiver_address] += ', '.join(selected)

        # update dataframe
        df_sent.at[index, 'sent_files'] = updated_sent

        # save df
        df_sent.to_csv('sent_files.csv', header=True, index=False)

        # open attachment
        for filename in selected:
            attachment = open(files_dir / filename, 'rb')
            
            # instance of MIMEBase and named as p 
            p = MIMEBase('application', 'octet-stream') 
            
            # To change the payload into encoded form 
            p.set_payload((attachment).read()) 

            # encode into base64 
            encoders.encode_base64(p) 

            p.add_header('Content-Disposition', "attachment; filename= %s" % filename) 

            # attach the instance 'p' to instance 'msg' 
            msg.attach(p) 

        # session = server, parameters: (location, port)
        s = smtplib.SMTP('smtp.gmail.com', 587)

        # TLS mode (transport layer security) to encrypt credentials
        s.starttls()

        # authentication
        s.login(sender_address, sender_password)

        # message body
        text = msg.as_string()

        # send email
        s.sendmail(sender_address, receiver_address, text)

        # terminate session
        s.quit()


if __name__ == '__main__':
    main()
