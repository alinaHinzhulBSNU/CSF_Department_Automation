import os
import pandas as pd
from docxtpl import DocxTemplate

# turn off chained assignment warning
pd.options.mode.chained_assignment = None

# ECTS
def score_to_ects(score):
    if 90 <= score <= 100:
        return "A"
    elif 82 <= score < 90:
        return "B"
    elif 75 <= score < 82:
        return "C"
    elif 64 <= score < 75:
        return "D"
    elif 60 <= score < 64:
        return "E"
    elif 35 <= score < 60:
        return "FX"
    elif score < 35:
        return "F"
    else:
        return None

# НАЦІОНАЛЬНА ШКАЛА  
def score_to_national_grade(ects):
    if ects == "A":
        return "Відмінно/Excellent"
    elif ects == "B" or ects == "C":
        return "Добре/Good"
    elif ects == "D" or ects == "E":
        return "Задовільно/Satisfactory"
    else:
        return None
    
# Перелік навчальних груп, які випускаються
specialties = {
    "АКІТ (магістри)": [671],
    "АКІТ": [471],
    "ІПЗ": [408, 409],
    "КІ": [405],
    "КН": [401, 402]
}

if __name__ == "__main__":
    # DATA SOURCES
    
    # Розподіл по групам
    graduates_codes = pd.read_excel("data/graduates_codes.xlsx")

    # Персональні дані здобувачів (дата народження, атестат, тощо)
    graduates = pd.read_excel("data/graduates.xlsx")

    # Дані про оцінки
    marks = pd.read_csv("data/marks.csv")

    # ДЛЯ КОЖНОЇ СПЕЦІАЛЬНОСТІ
    for specialty in specialties:

        # TEMPLATE
        template = DocxTemplate(f"templates/{specialty}.docx")

        groups = specialties[specialty]

        # ДЛЯ КОЖНОЇ ГРУПИ
        for group in groups:
            group_graduates_codes = list(graduates_codes[graduates_codes["group"] == group]["student_code"])

            # ДЛЯ КОЖНОГО ЗДОБУВАЧА
            for graduate_code in group_graduates_codes:

                # Персональні дані та дані про оцінки
                graduate_data = graduates[graduates["code"] == graduate_code]
                graduate_marks = marks[marks["code"] == graduate_code]

                # Впорядкування оцінок за категоріями
                theoretical_mask = ~(graduate_marks["type"].isin({"ПР", "КП", "КР"}) | graduate_marks["subject"].str.contains("ДОП"))
                theoretical = [row for index, row in graduate_marks[theoretical_mask].iterrows()]

                course_works_mask = graduate_marks["type"].isin({"КР", "КП"})
                course_works = [row for index, row in graduate_marks[course_works_mask].iterrows()]

                practices_mask = graduate_marks["type"].isin({"ПР"})
                practices = [row for index, row in graduate_marks[practices_mask].iterrows()]

                additional_mask = graduate_marks["subject"].str.contains("ДОП")
                additional = [row for index, row in graduate_marks[additional_mask].iterrows()]

                # ДОПи (сума кредитів та годин)
                additional_credits_sum = graduate_marks[additional_mask]["credits"].sum()
                additional_hours_sum = int(graduate_marks[additional_mask]["hours"].sum())
                additional_credits_hours = f"{additional_credits_sum} ({additional_hours_sum})"

                # Форматування дат
                graduate_data["birth_date"] = pd.to_datetime(graduate_data["birth_date"])
                graduate_data["birth_date"] = graduate_data["birth_date"].dt.strftime('%d/%m/%Y')

                graduate_data["study_start"] = pd.to_datetime(graduate_data["study_start"])
                graduate_data["study_start"] = graduate_data["study_start"].dt.strftime('%d/%m/%Y')

                graduate_data["study_end"] = pd.to_datetime(graduate_data["study_end"])
                graduate_data["study_end"] = graduate_data["study_end"].dt.strftime('%d/%m/%Y')

                # Середній бал (БЕЗ урахування диплому)
                mean_grade = graduate_marks[graduate_marks["grade"] > 59]["grade"].mean() # не враховувати борги (0) при розрахунку середнього бала
                mean_grade_ects = score_to_ects(mean_grade)
                mean_grade_national = score_to_national_grade(mean_grade_ects)

                # Дані, якими буде заповнено шаблон
                context = {
                    "surname_ukr": graduate_data["surname_ukr"].values[0],
                    "surname_eng": graduate_data["surname_eng"].values[0],
                    "name_ukr": graduate_data["name_ukr"].values[0],
                    "name_eng": graduate_data["name_eng"].values[0],
                    "card": graduate_data["card_id"].values[0],
                    "birth_date": graduate_data["birth_date"].values[0],
                    "study_start": graduate_data["study_start"].values[0],
                    "study_end": graduate_data["study_end"].values[0],
                    "certificate": graduate_data["school_certificate"].values[0],
                    "honours_ukr": graduate_data["honours_ukr"].values[0],
                    "honours_eng": graduate_data["honours_eng"].values[0],
                    "theoretical": theoretical,
                    "course_works": course_works,
                    "practices": practices,
                    "additional": additional,
                    "mean_grade": mean_grade,
                    "mean_grade_ects": mean_grade_ects,
                    "mean_grade_national": mean_grade_national,
                    "additional_credits_hours": additional_credits_hours
                }
                
                # Рендеринг шаблону
                template.render(context)
                
                # Збереження файлу
                file_name = f"result/{specialty}/{group}/{graduate_data['surname_ukr'].values[0]}.docx"
                os.makedirs(os.path.dirname(file_name), exist_ok=True)
                template.save(file_name)
