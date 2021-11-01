import pandas as pd
bio = pd.read_excel('bio.xlsx', header=None)

def make_html(fio: str, course: str, prev_course: str, work: str, educ: str, ph: bool, photo: str, output: str):
    """
    Создаём метод, который создаёт html-страницу с биографией преподавателя
    Аргументы:
    fio - ФИО преподавателя
    course - Название читаемого курса,
    prev_course - Курсы, прочитанные в ЭМШ в прошлом (год, название, сопреподаватели и стажёры, классы, направление),
    work - Место учёбы (работы),
    educ - Образование,
    ph - есть ли фото,
    photo - id фото на гугл-диске,
    output - название html файла, который получится на выходе
    """
    if ph:
        text = f"""
  <html>
 <head>
  <style type="text/css">
   TABLE {{
    width: 500px;
    border-collapse: collapse;
   }}
   TD, TH {{
    padding: 3px; 
    border: 1px solid black;
   }}
   TH {{
    background: #87325d;
   }}
     table {{
font-family: "Lucida Sans Unicode", "Lucida Grande", Sans-Serif;
font-size: 14px;
border-collapse: collapse;
text-align: center;
}}
th, td:first-child {{
background: #a15069;
color: white;
padding: 10px 20px;
}}
th, td {{
border-style: solid;
border-width: 0 1px 1px 0;
border-color: white;
}}
td {{
background: #fcdce6;
}}
th:first-child, td:first-child {{
text-align: left;
}}
  </style>

 </head>

 <body>
 <table>
<tbody>
<tr style="height: 24px;">
<td style="height: 24px; width: 39.2857%;" width="281"><strong>ФИО</strong></td>
<td style="height: 24px; width: 58.6905%;" width="429">{fio}</td>
</tr>
<tr style="height: 24px;">
<td style="height: 24px; width: 39.2857%;" width="281"><strong>Название читаемого курса</strong></td>
<td style="height: 24px; width: 58.6905%;" width="429">{course}</td>
</tr>
<tr style="height: 66px;">
<td style="height: 66px; width: 39.2857%;" width="281"><strong>Курсы, прочитанные в ЭМШ в прошлом (год, название, сопреподаватели и стажёры, классы, направление)</strong></td>
<td style="height: 66px; width: 58.6905%;" width="429">{prev_course}</td>
</tr>
<tr style="height: 24px;">
<td style="height: 24px; width: 39.2857%;" width="281"><strong>Место учёбы (работы)</strong></td>
<td style="height: 24px; width: 58.6905%;" width="429">{work}</td>
</tr>
<tr style="height: 24px;">
<td style="height: 24px; width: 39.2857%;" width="281"><strong>Образование</strong></td>
<td style="height: 24px; width: 58.6905%;" width="429">{educ}</td>
</tr>
<tr style="height: 45px;">
<td style="height: 45px; width: 39.2857%;" width="281"><strong>Фотография преподавателя :)</strong></td>
<td style="height: 45px; width: 58.6905%;text-align: center;" width="429"><p class="aligncenter"><img src="https://drive.google.com/uc?export=view&id={photo}" width=180px></p></a></td>
</tr>
</tbody>
</table>
</body>
</html>
"""
    else:
        text = f"""
  <html>
 <head>
 <style type="text/css">
   TABLE {{
    width: 1000px;
    border-collapse: collapse;
   }}
   TD, TH {{
    padding: 3px; 
    border: 1px solid black;
   }}
   TH {{
    background: #87325d;
   }}
   table {{
font-family: "Lucida Sans Unicode", "Lucida Grande", Sans-Serif;
font-size: 14px;
border-collapse: collapse;
text-align: center;
}}
th, td:first-child {{
background: #a15069;
color: white;
padding: 10px 20px;
}}
th, td {{
border-style: solid;
border-width: 0 1px 1px 0;
border-color: white;
}}
td {{
background: #fcdce6;
}}
th:first-child, td:first-child {{
text-align: left;
}}
  </style>
 </head>
 <body>
 <table>
<tbody>
<tr style="height: 24px;">
<td style="height: 24px; width: 39.2857%;" width="281"><strong>ФИО</strong></td>
<td style="height: 24px; width: 58.6905%;" width="629">{fio}</td>
</tr>
<tr style="height: 24px;">
<td style="height: 24px; width: 39.2857%;" width="281"><strong>Название читаемого курса</strong></td>
<td style="height: 24px; width: 58.6905%;" width="429">{course}</td>
</tr>
<tr style="height: 66px;">
<td style="height: 66px; width: 39.2857%;" width="281"><strong>Курсы, прочитанные в ЭМШ в прошлом (год, название, сопреподаватели и стажёры, классы, направление)</strong></td>
<td style="height: 66px; width: 58.6905%;" width="429">{prev_course}</td>
</tr>
<tr style="height: 24px;">
<td style="height: 24px; width: 39.2857%;" width="281"><strong>Место учёбы (работы)</strong></td>
<td style="height: 24px; width: 58.6905%;" width="429">{work}</td>
</tr>
<tr style="height: 24px;">
<td style="height: 24px; width: 39.2857%;" width="281"><strong>Образование</strong></td>
<td style="height: 24px; width: 58.6905%;" width="429">{educ}</td>
</tr>
</tbody>
</table>
</body>
</html>
"""
    file = open(f"{output}","w")
    file.write(text)
    file.close()

teachers = int((len(bio)+1)/7)
for i in range(teachers):
    fio, course, prev_course, work, educ, photo = bio[1][7*i], bio[1][7*i+1], bio[1][7*i+2], bio[1][7*i+3], bio[1][7*i+4], bio[1][7*i+5]
    try:
        make_html(fio, course, prev_course, work, educ, True, photo[33:], f'био/{fio}.html')
    except:
        make_html(fio, course, str(prev_course), work, educ, False, photo, f'био/{fio}.html')