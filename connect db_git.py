import pandas as pd
import pyodbc
# запуск приложений, в частности OutLook
import win32com.client as win32
from tabulate import tabulate

#функция для формирования сообщения
def send_email(to_email, subject, text, path_template):
    outlook = win32.Dispatch('outlook.application')
    mail = outlook.CreateItemFromTemplate(path_template)
    account = None
    for acc in mail.Session.Accounts:
        if acc.SmtpAddress == "от кого":
            account = acc
            break
    mail._oleobj_.Invoke(*(64209, 0, 8, 0, account))
    mail.To = to_email
    mail.HTMLBody = text.to_html() + mail.HTMLBody
    mail.Subject = subject
    mail.Send()

blps_dmt_oilrf = pyodbc.connect('Driver={SQL Server};''сервер;''Database=база данных;''Trusted_Connection=yes;')

df = pd.read_sql_query("""
                            SELECT MAX([_load_date]) as [last_date], 'view1' as [view_name] FROM view1
                            UNION
                            SELECT MAX([_load_date]) as [last_date], 'view2' as [view_name] FROM view2
                            UNION
                            and etc
                        """
, blps_dmt_oilrf)

send_email('кому', 'тема', df, r"path")
