import imaplib
import email
import time
from email.header import decode_header
import os
from fpdf import FPDF
import eel
from collections import Counter
import matplotlib.pyplot as plt



def matplotlibMakingGraph(imap, username, password):
    a = []
    status, messages = imap.select("INBOX")
    # total number of emails
    messages = int(messages[0])
    tea = messages
    for i in range(messages, messages - (tea - 2), -1):
        # fetch the email message by ID
        res, msg = imap.fetch(str(i), "(RFC822)")
        for response in msg:
            if isinstance(response, tuple):
                # parse a bytes email into a message object
                msg = email.message_from_bytes(response[1])

                Date, encoding = decode_header(msg["Date"])[0]
                if isinstance(Date, bytes):
                    Date = Date.decode(encoding)

                abc = Date[0:17]
                a.append(abc)

    v = Counter(a)
    q = dict(v)
    print(q)
    plt.bar(*zip(*q.items()))
    ml = os.getcwd()
    cw = os.path.join(os.getcwd(), "web")
    os.chdir(cw)
    if os.path.exists("example.png"):
        os.remove("example.png")
        plt.savefig("example.png")
        os.chdir(ml)
    else:
        plt.savefig("example.png")
        os.chdir(ml)


def firstMailExtraction(imap):
    body=""
    z = open("omkar.txt", "w")
    status, messages = imap.select("INBOX")
    # number of top emails to fetch
    N = 1
    # total number of emails
    messages = int(messages[0])
    for i in range(messages, messages - N, -1):
        # fetch the email message by ID
        res, msg = imap.fetch(str(i), "(RFC822)")
        for response in msg:
            if isinstance(response, tuple):
                # parse a bytes email into a message object
                msg = email.message_from_bytes(response[1])
                # decode the email subject
                subject, encoding = decode_header(msg["Subject"])[0]
                if isinstance(subject, bytes):
                    # if it's a bytes, decode to str
                    subject = subject.decode(encoding)
                # decode email sender
                From, encoding = decode_header(msg.get("From"))[0]
                if isinstance(From, bytes):
                    From = From.decode(encoding)

                # print("Subject:", subject)
                z.write("subject = ")
                z.write(subject)
                # print("From:", From)
                z.write("\nfrom = ")
                z.write(From)
                # if the email message is multipart
                if msg.is_multipart():

                    # iterate over email parts
                    for part in msg.walk():
                        # extract content type of email
                        content_type = part.get_content_type()
                        content_disposition = str(part.get("Content-Disposition"))
                        try:
                            # get the email body
                            body = part.get_payload(decode=True).decode()

                        except:
                            pass

                        if content_type == "text/plain" and "attachment" not in content_disposition:
                            # print text/plain emails and skip attachments
                            z.write("\nbody = ")
                            z.write(body)
                            # print(body)

                        elif "attachment" in content_disposition:
                            # download attachment
                            filename = part.get_filename()
                            if filename:
                                folder_name = clean(subject)
                                if not os.path.isdir(folder_name):
                                    # make a folder for this email (named after the subject)
                                    os.mkdir(folder_name)
                                filepath = os.path.join(folder_name, filename)
                                # download attachment and save it
                                open(filepath, "wb").write(part.get_payload(decode=True))
                else:
                    # extract content type of email
                    content_type = msg.get_content_type()
                    # get the email body
                    body = msg.get_payload(decode=True).decode()
                    if content_type == "text/plain":
                        # print only text email parts
                        print(body)

                print("=" * 100)
    # close the connection and logout
    z.close()

    pdf = FPDF()

    # Add a page
    pdf.add_page()

    # set style and size of font
    # that you want in the pdf
    pdf.set_font("Arial", size=15)

    # open the text file in read mode
    f = open("omkar.txt", "r")

    # insert the texts in pdf
    for x in f:
        pdf.cell(200, 10, txt=x, ln=1, align='C')

    # save the pdf with name .pdf
    pdf.output("omkarpdf.pdf")
    print(".....conversion done successfully......")
    imap.close()
    imap.logout()


def emailvalidation():
    try:
        username = input("enter email address")
        password = input("enter password")
        # create an IMAP4 class with SSL
        imap = imaplib.IMAP4_SSL("imap.gmail.com")
        # authenticate
        imap.login(username, password)
        matplotlibMakingGraph(imap, username, password)
        firstMailExtraction(imap)
        return (imap)
    except Exception as e:
        print("invalid email or password please enter again")
        emailvalidation()


def clean(text):
    # clean text for creating a folder
    return "".join(c if c.isalnum() else "_" for c in text)


emailvalidation()
time.sleep(1)

eel.init("web")
eel.expose()


def getdata(data, data1):
    if data == "showstatistics" and data1 == "1234":
        return "ok"


eel.start('index.html', size=(800, 700))

#   aitglobal105@gmail.com
#   Aitglobal@105

