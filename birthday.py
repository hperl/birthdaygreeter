from string import Template
import smtplib
import email
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.header import Header
import unittest
from datetime import datetime, date
import os

import xlrd


class Person(object):
    def __init__(self, firstname, lastname, birthday, sex, email):
        self.firstname = firstname
        self.lastname = lastname
        self.birthday = datetime.strptime(birthday, "%Y-%m-%d")
        self.email = email
        self.__sex = sex

    def is_male(self):
        return self.__sex == 'm'

    def is_female(self):
        return not self.is_male()

    @property
    def has_birthday(self):
        today = date.today()
        return self.birthday.month == today.month and self.birthday.day == today.day


class PeopleParser(object):
    def __init__(self, path):
        self.__path = path

    def parse(self):
        self.__sheet = xlrd.open_workbook(self.__path).sheets()[0]
        return self.__parse()

    def __parse(self):
        self.__parse_headers()
        return [self.__parse_row(i) for i in range(1, self.__sheet.nrows)]

    def __parse_headers(self):
        first_row = self.__sheet.row(0)
        self.__headers = {}
        for idx, cell in enumerate(first_row):
            self.__headers[cell.value] = idx

    def __parse_row(self, idx):
        row = self.__sheet.row(idx)
        return Person(
            firstname=row[self.__headers['Vorname']].value,
            lastname=row[self.__headers['Nachname']].value,
            birthday=row[self.__headers['Geburtsdatum']].value,
            email=row[self.__headers['E-Mail']].value,
            sex=row[self.__headers['Geschlecht']].value
        )

class BirthdayMail(object):
    def __init__(self, person):
        self.person = person

    def __text(self):
        if self.person.is_male():
            anrede = "Lieber %s" % self.person.firstname
        else:
            anrede = "Liebe %s" % self.person.firstname
        with open('template.txt', 'r') as t:
            return Template(t.read()).substitute(anrede=anrede)

    def __message(self):
        msg = MIMEMultipart()
        msg['Subject'] = Header("Herzlichen Glückwunsch zum Geburtstag!", "utf-8")
        msg['From'] = "geburtstag@yfu.de"
        msg['To'] = self.person.email
        msg['Date'] = email.utils.formatdate()
        msg.attach(MIMEText(self.__text(), "plain", "utf-8"))
        return msg

    def send(self):
        msg = self.__message()
        smtp = smtplib.SMTP('192.168.100.3')
        smtp.sendmail("henning.perl@yfu-deutschland.de", self.person.email, msg.as_string())
        print("<%s> sent." % self.person.email)
        smtp.quit()


def main():
    for person in PeopleParser("Mitglieder.xls").parse():
        if person.has_birthday:
            BirthdayMail(person).send()


class TestBirthayPeople(unittest.TestCase):
    def test_new_person(self):
        p = Person("Kai", "Uwe", "1957-01-02", "m", "")
        self.assertEqual(p.birthday.year, 1957)
        self.assertTrue(p.is_male())

    def test_has_birthday(self):
        today = date.today().strftime("%Y-%m-%d")
        p = Person("Kai", "Uwe", today, "m", "")
        self.assertTrue(p.has_birthday)

    def test_people_parser(self):
        for person in PeopleParser("test.xls").parse():
            self.assertEqual(person.lastname, "Nachname")
            self.assertEqual(person.email, "test@mail.de")


if __name__ == '__main__':
    if os.environ.get('TEST'):
        unittest.main()
    else:
        main()

