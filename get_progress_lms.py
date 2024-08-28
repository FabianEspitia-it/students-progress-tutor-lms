import io
import pandas as pd
import requests
import pymysql
import os
import datetime

from requests.auth import HTTPBasicAuth
from dotenv import load_dotenv


load_dotenv()


def create_connection():
    connection = pymysql.connect(
        host=os.getenv("DB_HOST"),
        user=os.getenv("DB_USER"),
        password=os.getenv("DB_PASSWORD"),
        database=os.getenv("DB_NAME")
    )
    return connection


def get_modules_ids() -> list:

    response = requests.get(os.getenv("GET_MODULES_IDS_URL"), auth=HTTPBasicAuth(os.getenv("USER_NAME"), os.getenv("PASSWORD")))

    if response.status_code == 200:
        json_data = response.json()
        return [item['ID'] for item in json_data['data']]
    
    return response.status_code


def get_total_amount_classes() -> int:

    modules_ids: list = get_modules_ids()

    total_classes: int = 0

    for module_id in modules_ids:
        response = requests.get(f"{os.getenv("GET_LESSONS_BY_MODULE_URL")}{module_id}", auth=HTTPBasicAuth(os.getenv("USER_NAME"), os.getenv("PASSWORD")))
        if response.status_code == 200:
            json_data = response.json()
            total_classes += len(json_data['data'])
        else:
            print(response.status_code)
            break

    return total_classes


def get_students_ids() -> list:

    connection = create_connection()
    try:
        with connection.cursor() as cursor:
            cursor.execute(os.getenv("GET_STUDENTS_IDS_QUERY"))
            result = cursor.fetchall()
            return [item[0] for item in result]

    finally:
        connection.close()


def get_progress_student():
    
    students_ids: list = get_students_ids()
    students_progress: list[dict] = []

    emails: list[str] = ['team@makers.ngo',
                         'test@grayola.io',
                         'samuel@grayola.io',
                         'tutor@makers.ngo'
                         ]

    counter: int = len(students_ids) - len(emails)

    


    connection = create_connection()
    try: 
        for student_id in students_ids:
            
                with connection.cursor() as cursor:
                    cursor.execute(f"{os.getenv("GET_STUDENT_CLASSES_SEEN_QUERY")}{student_id} {os.getenv("GET_STUDENT_CLASSES_SEEN_QUERY_TWO")}")
                    result = cursor.fetchone()
                    total_classes = get_total_amount_classes()

                    progress = str(round((result[0] / total_classes), 2) * 100) 

                    cursor.execute(f"{os.getenv("GET_STUDENT_EMAIL_QUERY")}{student_id};")
                    email = cursor.fetchone()[0]


                    
                    if email not in emails:
                        
                        students_progress.append({
                            "email": email,
                            "progress": f"{progress}%"
                        })

                        print(f"Student Progress: {email} - {progress}%")
                        print(f"Students Left: {counter}")
                        counter -= 1        
    finally:
        connection.close()

    
    return students_progress


def build_excel() -> io.BytesIO:
   
    students_progress = get_progress_student()

    df = pd.DataFrame(students_progress)
    df['date'] = datetime.datetime.now()

    
    buffer = io.BytesIO()

  
    with pd.ExcelWriter(buffer, engine='openpyxl') as writer:

        df.to_excel(writer, index=False, header=True)

        worksheet = writer.sheets['Sheet1']

        for col in worksheet.columns:
                max_length = 0
                column = col[0].column_letter
                for cell in col:
                    try:
                        if len(str(cell.value)) > max_length:
                            max_length = len(cell.value)
                    except:
                        pass
                adjusted_width = (max_length + 2)
                worksheet.column_dimensions[column].width = adjusted_width  

    
    buffer.seek(0)

    return buffer


def save_to_desktop() -> None:
   
    buffer = build_excel()
    
 
    desktop_path = os.path.join(os.path.expanduser("~"), "onedrive", "escritorio")
    file_path = os.path.join(desktop_path, "students_progress.xlsx")

   
    with open(file_path, 'wb') as f:
        f.write(buffer.getvalue())

    print(f"Excel saved")



save_to_desktop()




