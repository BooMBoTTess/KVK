import win32com.client as win32
import pandas as pd
import datetime as dt

import xlwings as xw


def send_message_heads(name, email, department, deadline, left_days, number_kvk):
    outlook = win32.Dispatch('outlook.application')
    mail = outlook.CreateItem(1)

    mail.Recipients.Add(email)
    mail.Recipients.ResolveAll()
    mail.ReminderSet = True
    mail.ReminderMinutesBeforeStart = 2
    mail.MeetingStatus = 1

    mail.start = dt.datetime.strptime(deadline, '%d.%m.%Y')



    mail.Subject = 'Сдача карты внутреннего контроля'
    if left_days < 0:
        #Дедлайн ушел минус плохо все
        mail.Body = f"Уважаемый (-ая) {name}\n" \
                    f"Вам нужно сдать карту контроля, в связи с изменением в штате.\n" \
                    f"Срок сдачи до: {deadline}. Вы просрочили сдачу на: {left_days * -1} дней.\n" \
                    f"Вам нужно сдать {number_kvk} карт контроля."
    else:
        mail.Body = f"Уважаемый (-ая) {name}\n" \
                    f"Вам нужно сдать карту контроля, в связи с изменением в штате.\n" \
                    f"Срок сдачи до: {deadline}. Осталось {left_days} дней.\n" \
                    f"Вам нужно сдать {number_kvk} карт контроля."
    mail.Send()

def head_dep_message(df):
    app = xw.App(visible=False)
    wb = xw.Book('kadrifile.xlsx')
    wb.save('kadrifile.xlsx')
    wb.close()
    app.quit()

    target_cards_dataframe = df[df['Осталось дней'].notna()]
    df_names_emails = pd.read_excel('kadrifile.xlsx', sheet_name='Историярасслылок2', skiprows=3, usecols='J:M')

    for index, row in target_cards_dataframe.iterrows():
        name = row['Начальник отдела']
        email = df_names_emails[df_names_emails['Отдел'] == row['Отдел']]['Почта.1'].to_string(index=False)
        department = row['Отдел']
        deadline = row['Срок сдачи']
        left_days = row['Осталось дней']
        number_kvk = row['Количество карт контроля']
        send_message_heads(name, email, department, deadline, left_days, number_kvk)

def send_message_kurator(name, kur_email, deps, head_names: list, deadlines: list, left_days: list):
    outlook = win32.Dispatch('outlook.application')
    mail = outlook.CreateItem(0)
    mail.To = kur_email
    mail.Subject = 'Сдача карты внутреннего контроля'
    mail.Body = f'Уважаемый (-ая) {name}.\n' \
                f'Вашим отделам требуется сдать карты контроля:\n'
    for i in range(len(deps)):
        mail.Body += f'{deps[i]}. Глава отдела {head_names[i]}. Срок до {deadlines[i]}. Осталось дней: {left_days[i]}\n'
    mail.Send()


def kurator_message(df):
    app = xw.App(visible=False)
    wb = xw.Book('kadrifile.xlsx')
    wb.save('kadrifile.xlsx')
    wb.close()
    app.quit()

    target_cards_dataframe = df[df['Осталось дней'].notna()]
    df_kurator_emails = pd.read_excel('kadrifile.xlsx', sheet_name='Историярасслылок2', skiprows=3, usecols='N:O')


    for index, row in df_kurator_emails.drop_duplicates().iterrows():
        name = row['Куратор_1']
        email = row['Почта_1']
        if target_cards_dataframe[target_cards_dataframe['Куратор'] == name].shape[0] != 0:
            df_under_kurator = target_cards_dataframe[target_cards_dataframe['Куратор'] == name]
            deps = df_under_kurator['Отдел'].tolist()
            head_names = df_under_kurator['Начальник отдела'].tolist()
            deadlines = df_under_kurator['Срок сдачи'].tolist()
            left_days = df_under_kurator['Осталось дней'].tolist()
            send_message_kurator(name, email, deps, head_names, deadlines, left_days)