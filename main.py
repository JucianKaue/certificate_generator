import datetime

import openpyxl

from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
import os

import glob

pastaApp = os.path.dirname(__file__)


# Organiza as informações sobre o curso.
class Course:
    def __init__(self, name, date, workload, contents):
        date = [date['start'].split('/'), date['end'].split('/')]
        self.name_ = name
        self.date_ = [datetime.date(day=int(date[0][0]), month=int(date[0][1]), year=int(date[0][2])),
                      datetime.date(day=int(date[1][0]), month=int(date[1][1]), year=int(date[1][2]))]
        self.workload_ = workload
        self.contents = contents


# Organiza as informações e métodos relacionados ao certificado
class Certificate:
    # Salva as informações importantes para o certificado, sendo elas,  aluno, curso, data e o texto.
    def __init__(self, student, course):
        self.student = student
        self.course = course
        self.date = datetime.datetime.today()
        self.text_ = f"Certificamos que {self.student} participou com êxito do evento {self.course.name_},  " \
                     f"realizado de {Certificate.convert_date(self.course.date_[0])} a " \
                     f"{Certificate.convert_date(self.course.date_[1])} de forma virtual, " \
                     f"contabilizando carga horária total de {self.course.workload_} horas."
        self.text_authentication =  "Para verificar a autenticidade deste documento entre em" \
                                    " https://sig.ifc.edu.br/documentos/ informando seu número: " \
                                    "161, ano: 2021, tipo: DECLARAÇÃO, data de emissão: 06/11/2021 " \
                                    "e o código de verificação: d01d22c977"

    # Método responsável por criar o certificado.
    def generate_certification(self):
        Certificate.draw_text('CERTIFICADO', size=50, position=[600 / 2, 750], color='#13C3AF')     # Define o título

        text = Certificate.split_text(self.text_, 40)   # Chama a função split_text (linha 65)
        pos_Y = 690                                     # Define  a posição inicial horizontal da primeira linha

        # Escreve o texto
        for i in range(len(text)):
            Certificate.draw_text(text[i], position=[600 / 2, pos_Y], size=15)
            pos_Y -= 20

        pos_Y -= 45
        Certificate.draw_text('COMPOSIÇÃO DO CURSO', size=20, position=[600 / 2, pos_Y])
        pos_Y -= 30
        for i in range(len(self.course.contents)):
            Certificate.draw_text(self.course.contents[i], position=[600 / 2, pos_Y], size=13)
            pos_Y -= 20

        pos_Y -= 185
        canva.drawImage('images/logo.jpeg', 200, pos_Y)

        pos_Y -= 110
        canva.drawImage('images/fabão.png', 50, pos_Y, 150, 80)
        canva.drawImage('images/rafa.png', 400, pos_Y, 150, 80)
        canva.drawImage('images/IFC.png', 240, pos_Y-20, 120, 120)

        p = 60
        for sentence in Certificate.split_text(self.text_authentication, 95):
            Certificate.draw_text(sentence, position=[600 / 2, p], size=9)
            p -= 10

        Certificate.draw_text(Certificate.convert_date(self.date), position=[600 / 2, 20], size=12)

    @staticmethod
    # Recebe uma string, divide ela a cada determinado numero de caracteres e retorna uma lista com as strings resultantes.
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
    # Responsável por converter datas do formato datetime para a data escrita por extenso.
    def convert_date(date):
        months = [ 'Janeiro', 'Fevereiro', 'Março', 'Abril', 'Maio', 'Junho', 'Julho', 'Agosto', 'Setembro', 'Outubro', 'Novembro', 'Dezembro']
        return f"{date.day} de {months[(date.month-1)]} de {date.year}"

    @staticmethod
    # Responsável por escrever os textos no certificado.
    def draw_text(text, position, size, color='black'):
        canva.setFont("Helvetica", size)
        canva.setFillColor(color)
        canva.drawCentredString(position[0], position[1], text)


# Responsável por organizar e coletar os dados das tabelas.
class Table:
    def __init__(self, arquive_name, table_name=''):
        self.arquive_ = openpyxl.load_workbook(filename=f"{arquive_name}", read_only=True)
        if table_name != '':
            self.table = self.arquive_[table_name]

    # Recebe os titulos das colunas e retorna uma lista contendo dicionários que representam as linhas da tabela
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
    # Recebe o email do aluno e a lista de tabelas da chamada e retorna a porcentagem de presença
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

# Função responsável por enviar o e-mail
def send_email(mail_from, mail_to, certificate_name):
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
    message.set_content('Olá, gostariamos de lhe parabenizar pela conclusão do Curso '
                        '"Ensinando Arduino, Ciclo de Aulas On-Line de Eletrônica".\n\n'
                        'Segue em anexo o seu certificado de conclusão do curso')

    file = open(f'certificates\{certificate_name}', 'rb')
    file_data = file.read()
    file_name = certificate_name

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
              'Semáforo, projeto de semáforo com Arduino',
              'Controle de luminosidade, regulando a tensão com o potenciômetro',
              'Semáforo interativo, projeto de semáforo com pedestres',
              'Servo motor, controlando o servo motor']
)

# Abre a tabela de inscrição e salva o nome e o e-mail dos participantes.
table_subscription = Table(arquive_name="input\subscribed.xlsx", table_name="subscribed")
list_students = table_subscription.get_fields(fields_name={"Email": "Email Address",
                                           "First Name": "First Name",
                                           "Last Name": "Last Name",})

print(list_students)
# Descobre o nome e abre as tabelas que estivarem dentro de "input\presence"
list_tables_attendance = []
for i in glob.glob("*input\presence\*"):
    table = Table(arquive_name=f"{i}")
    table.table = table.arquive_[table.arquive_.sheetnames[0]]
    list_tables_attendance.append(table)

generate_amount = 0
not_generate_amount = 0

pastaApp = os.path.dirname(__file__)
for student in list_students:
    # Descobre a porcentagem de presença
    percentage_presence = Table.verify_presence(student={'e-mail': f'{student["Email Address"]}'},
                                                attendence_lists=list_tables_attendance)
    # Se a porcentagem de presença for maior ou igual a 70
    if percentage_presence[1] >= 70:
        # Gera o certificado
        canva = canvas.Canvas(pastaApp + f"\\certificates\certificado-{student['First Name']} {student['Last Name']}.pdf", pagesize=A4)
        certification = Certificate(student=f"{student['First Name'].upper()} {student['Last Name'].upper()}",
                                    course=course)
        certification.generate_certification()
        canva.save()
        # Mostra uma mensagem falando que deu tudo certo
        print(f"\033[1;32mO certificado de {student['First Name'].upper()} {student['Last Name'].upper()} foi gerado corretamente\033[0m\n"
              f"Presença: {percentage_presence[1]}")
        generate_amount += 1

        send_email(mail_from=['canal.top.arduino@gmail.com', 'zpmhbmumxrzdfsud'],
                   mail_to=f'{student["Email Address"].lower()}',
                   certificate_name=f"certificado-{student['First Name']} {student['Last Name']}.pdf")

    else:
        # Mostra uma mensagem avisando que o aluno não obteve frequências suficientes
        print(
            f"\033[1;31mO aluno(a) {student['First Name'].upper()} {student['Last Name'].upper()} não atingiu frequência suficiente\033[0m\n"
            f"Presença: {percentage_presence[1]}")
        not_generate_amount += 1
print(f"Gerados: {generate_amount}")
print(f"Não gerados: {not_generate_amount}")
