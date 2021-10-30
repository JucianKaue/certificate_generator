import datetime
from time import sleep

import openpyxl

from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
import os

import glob

pastaApp = os.path.dirname(__file__)


class Course:
    def __init__(self, name, date, workload, contents):
        date = [date['start'].split('/'), date['end'].split('/')]
        self.name_ = name
        self.date_ = [datetime.date(day=int(date[0][0]), month=int(date[0][1]), year=int(date[0][2])),
                      datetime.date(day=int(date[1][0]), month=int(date[1][1]), year=int(date[1][2]))]
        self.workload_ = workload
        self.contents = contents


class Certification:
    def __init__(self, person, course):

        self.person = person
        self.course = course
        self.date = datetime.datetime.today()
        self.text_ = f"Certificamos que {self.person} participou com êxito do evento {self.course.name_},  " \
                     f"realizado de {Certification.convert_date(self.course.date_[0])} a " \
                     f"{Certification.convert_date(self.course.date_[1])} de forma virtual, " \
                     f"contabilizando carga horária total de {self.course.workload_} horas."

    def generate_certification(self):
        Certification.draw_text('CERTIFICADO', size=50, position=[600/2, 750], color='#13C3AF')

        text = Certification.split_text(self.text_, 40)
        pos_Y = 690
        for i in range(len(text)):
            Certification.draw_text(text[i], position=[600/2, pos_Y], size=15)
            pos_Y -= 20

        pos_Y -= 40
        Certification.draw_text('COMPOSIÇÃO DO CURSO', size=20, position=[600 / 2, pos_Y])
        pos_Y -= 30
        for i in range(len(self.course.contents)):
            Certification.draw_text(self.course.contents[i], position=[600/2, pos_Y], size=13)
            pos_Y -= 20

        pos_Y -= 185
        canva.drawImage('images/logo.jpeg', 200, pos_Y)

        pos_Y -= 110
        canva.drawImage('images/fabão.png', 50, pos_Y, 150, 80)
        canva.drawImage('images/rafa.png', 400, pos_Y, 150, 80)
        canva.drawImage('images/IFC.png', 240, pos_Y-20, 120, 120)

        Certification.draw_text(Certification.convert_date(self.date), position=[600 / 2, 20], size=10)

    @staticmethod
    def split_text(text, caracters_per_line):
        words = text.split()

        counter = 0
        sentence = ""
        text = []
        for word in words:
            counter += len(word)
            if counter > caracters_per_line or words[len(words)-1] == word:
                sentence += f"{word} "
                text.append(sentence)
                counter = 0
                sentence = ""
            else:
                sentence += f"{word} "

        return text

    @staticmethod
    def convert_date(date):
        months = [ 'Janeiro', 'Fevereiro', 'Março', 'Abril', 'Maio', 'Junho', 'Julho', 'Agosto', 'Setembro', 'Outubro', 'Novembro', 'Dezembro']
        return f"{date.day} de {months[(date.month-1)]} de {date.year}"

    @staticmethod
    def draw_text(text, position, size, color='black'):
        canva.setFont("Helvetica", size)
        canva.setFillColor(color)
        canva.drawCentredString(position[0], position[1], text)


class Table:
    def __init__(self, arquive_name, table_name=''):
        self.arquive_ = openpyxl.load_workbook(filename=f"{arquive_name}", read_only=True)
        if table_name != '':
            self.table = self.arquive_[table_name]

    def get_fields(self, fields_name):
        lines_list = []
        indexes = {}
        counter = 0
        for line in self.table:
            if counter == 0:
                counter_cell = 0
                for cell in line:
                    for field in fields_name.values():
                        if cell.value == field:
                            indexes[f'{field}'] = counter_cell
                    counter_cell += 1
            else:
                lines = {}
                for e in indexes.items():
                    lines[f'{e[0]}'] = line[e[1]].value
                lines_list.append(lines)
            counter += 1
        return lines_list

    @staticmethod
    def verify_presence(student, attendence_lists):
        list_all_attendences = []
        for table in attendence_lists:
            list_emails = []
            emails = table.get_fields({"Email": "Qual é o seu e-mail?"})
            for email in emails:
                list_emails.append(email['Qual é o seu e-mail?'].strip().lower())
            list_all_attendences.append(list_emails)

        attendences = 0
        for i in range(0, len(list_all_attendences)):
            if student['e-mail'].strip().lower() in list_all_attendences[i]:
                attendences += 1

        percent = int((attendences*100)/len(list_all_attendences))
        return student['e-mail'], percent


def send_email(mail_from, mail_to):
    import smtplib
    from email.message import EmailMessage

    smtp_ssl_host = 'smtp.gmail.com'
    smtp_ssl_port = 465

    username = mail_from[0]
    password = mail_from[1]

    mail_from = mail_from[0]
    mail_to = [mail_to]

    message = EmailMessage()
    message['Subject'] = "Certificado de Conclusão de Curso"
    message['From'] = mail_from
    message['To'] = mail_to
    message.set_content("HELLO WORLD!!!")

    file = open('certificado.pdf', 'rb')
    file_data = file.read()
    file_name = file.name

    message.add_attachment(file_data, maintype='pdf', subtype="pdf", filename=file_name)

    server = smtplib.SMTP_SSL(smtp_ssl_host, smtp_ssl_port)
    server.login(username, password)
    server.send_message(message)
    server.quit()


course = Course(
    name="Ensinando Arduino, Ciclo de Aulas On-Line de Eletrônica",
    date={'start': '12/08/2021', 'end': '23/09/2021'},
    workload="12",
    contents=['Apresentação do projeto',
              'Conhecendo o Arduino, noções de eletrônica e primeiro projeto',
              'Semáforo, projeto de semáfora com Arduino',
              'Controle de luminosidade, regulando a tensão com o potenciômetro',
              'Semáforo interativo, projeto de semáforo com pedestres',
              'Servo motor, controlando o servo motor']
)

# Abre a tabela de inscrição e salva o nome e o e-mail dos participantes.
table_subscription = Table(arquive_name="input\subscribed.xlsx", table_name="subscribed")
list_students = table_subscription.get_fields(fields_name={"Email": "Email Address",
                                           "First Name": "First Name",
                                           "Last Name": "Last Name",
                                           "Phone": "Phone Number"})

# Descobre o nome e abre as tabelas que estivarem dentro de "input\presence"
list_tables_attendance = []
for i in glob.glob("*input\presence\*"):
    table = Table(arquive_name=f"{i}")
    table.table = table.arquive_[table.arquive_.sheetnames[0]]
    list_tables_attendance.append(table)

generate_amount = 0
not_generate_amount = 0
students_passed = []

pastaApp = os.path.dirname(__file__)
for student in list_students:
    # Descobre a porcentagem de presença
    percentage_presence = Table.verify_presence(student={'e-mail': f'{student["Email Address"]}'},
                                                attendence_lists=list_tables_attendance)
    # Se a porcentagem for maior ou igual a 70
    if percentage_presence[1] >= 70:
        # Gera o certificado
        canva = canvas.Canvas(pastaApp + f"\\certificates\certificado-{student['First Name']} {student['Last Name']}.pdf", pagesize=A4)
        certification = Certification(person=f"{student['First Name'].upper()} {student['Last Name'].upper()}", course=course)
        certification.generate_certification()
        canva.save()
        # Mostra uma mensagem falando que deu tudo certo
        print(f"\033[1;32mO certificado de {student['First Name'].upper()} {student['Last Name'].upper()} foi gerado corretamente\033[0m\n"
              f"Presença: {percentage_presence[1]}")
        generate_amount += 1
    else:
        # Mostra uma mensagem avisando que o aluno não obteve frequências suficientes
        print(
            f"\033[1;31mO aluno(a) {student['First Name'].upper()} {student['Last Name'].upper()} não atingiu frequência suficiente\033[0m\n"
            f"Presença: {percentage_presence[1]}")
        students_passed.append(f"{student['First Name'].title()} {student['Last Name'].title()}")
        not_generate_amount += 1
    sleep(0.5)
print(f"Gerados: {generate_amount}")
print(f"Não gerados: {not_generate_amount}")

for s in students_passed:
    print(s)


send_email(mail_from=['', ''], mail_to='jucian_decezare2015@hotmail.com')
