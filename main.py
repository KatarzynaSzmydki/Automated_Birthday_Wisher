import win32com.client
import datetime as dt
import pandas as pd
import random
import os.path


# ---------------------------- CONSTANTS ------------------------------- #
emails = 'helenaszmydki@gmail.com'
word = []
meaning = []
word_choice = None
quote_choice = None
words_sent = []
quotes_sent = []


# ---------------------------- CONTENT ------------------------------- #

# Eng words
with open('./eng_words.txt', 'r') as f:
    eng_words = f.read()
    eng_words = eng_words.split('\n')
for i in range(len(eng_words)):
    if i % 2 == 0:
        word.append(eng_words[i].strip(' '))
    else:
        meaning.append(eng_words[i])

vocabulary = pd.DataFrame(list(zip(word, meaning)))
vocabulary.columns = ['word', 'meaning']


# Eng words sent
if os.path.exists('./eng_words_sent.csv'):
    words_sent = pd.read_csv('./eng_words_sent.csv')
    words_sent = list(words_sent['words_sent'])




# Quotes
with open('./quotes.txt', 'r') as f:
    quotes = f.read()
    quotes = quotes.split('\n')
    quotes = pd.DataFrame(quotes)
    quotes.columns = ['Quote']


# Quotes sent
if os.path.exists('./quotes_sent.csv'):
    quotes_sent = pd.read_csv('./quotes_sent.csv')
    quotes_sent = list(quotes_sent['quotes_sent'])




# ---------------------------- FUNCTIONS ------------------------------- #

def choose_word():

    global word_choice, words_sent

    word_choice = random.randint(0,len(vocabulary)-1)

    if word_choice in words_sent:
        choose_word()
    else:
        words_sent.append(word_choice)
        words_sent = pd.DataFrame(words_sent)
        words_sent.columns = ['words_sent']
        words_sent.to_csv('./eng_words_sent.csv')



def choose_quote():

    global quote_choice, quotes_sent

    quote_choice = random.randint(0,len(quotes)-1)

    if quote_choice in quotes_sent:
        choose_quote()
    else:
        quotes_sent.append(quote_choice)
        quotes_sent = pd.DataFrame(quotes_sent)
        quotes_sent.columns = ['quotes_sent']
        quotes_sent.to_csv('./quotes_sent.csv')






# ---------------------------- EMAIL SENDING ------------------------------- #


outlook = win32com.client.Dispatch('outlook.application')
mail = outlook.CreateItem(0)

mail.To = emails
mail.Subject = "Buckle up! It's your English o'clock!"

attachment1 = r'C:\Users\kszmydki\PycharmProjects\Automated_Birthday_Wisher\english.jfif'
attachment2 = r'C:\Users\kszmydki\PycharmProjects\Automated_Birthday_Wisher\learning.jfif'
ats = mail.Attachments
att1 = ats.Add(attachment1, 1, 0)
att2 = ats.Add(attachment2, 1, 0)

choose_quote()
choose_word()
Body1 = 'Quote of the day:\n\n'
Body2 = 'Word of the day:\n\n'
Quote = quotes.iloc[quote_choice].item()
Word = vocabulary['word'][word_choice]
Meaning = vocabulary['meaning'][word_choice]
mail.HTMLBody = '<img src="cid:{0:}" width=200 height=100> <br/><br/><b>{1:}</b>{2:}</b><br/><br/><b>{3:}</b><br/><br/>' \
                '<b>{4:}</b> - {5:} <br/><br/><img src="cid:{6:}" width=200 height=100> <br/>'.format(att1.FileName, Body1, Quote, Body2, Word, Meaning, att2.FileName)

mail.Send()



