import imaplib
import email
import re
import csv
from tokenz import mail_pass

def get_raw_data():
    '''Генератор, выдает бинарное письмо '''
    imap = imaplib.IMAP4_SSL('mail.stardustpaints.ru')
    imap.login('soloviev@stardustpaints.ru', mail_pass)
    imap.list()
    imap.select('INBOX')
    result, data = imap.search(None, '(SUBJECT "Request from [stardustpaints.ru]")')
    ids = data[0]
    ids_list = ids.split()
    for id_ in ids_list:
        result, data = imap.fetch(id_, "(RFC822)")
        yield data[0][1]

def get_mail_body(raw_data):
    '''Берем сырое письмо, возвращаем кортеж (дата, текст письма) '''
    mail = email.message_from_bytes(raw_data)
    body = mail.get_payload(decode=True)
    return mail['Date'], body.decode('utf-8')

def parse_data(date, body):
    '''Парсим данные из заявки, возвращает список'''
    patterns = [r'Name: (.+?)<br>', r'Phone: (.+?)<br>', r'Email: (.+?)<br>', r'Комментарий: (.+?)<br>']
    result = [date]
    for pattern in patterns:
        try:
            res = re.search(pattern, body).group(1)
        except:
            res = ''
        result.append(res)

    return result

def write_csv(all_requests):
    '''Пишем csv'''
    with open('from_site_reqs.csv', 'w') as f:
        writer = csv.writer(f)
        writer.writerows(all_requests)



def main():
    all_requests = [['Дата', 'Имя', 'Телефон', 'Почта', 'Комментарий']]
    raw_data_gen = get_raw_data()
    for raw_data in raw_data_gen:
        date, body = get_mail_body(raw_data)
        line = parse_data(date, body)
        all_requests.append(line)
    write_csv(all_requests)
        


if __name__ == "__main__":
    main()