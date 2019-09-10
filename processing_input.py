import datetime
import re
from Schedule import Schedule


def research(regular, string):
    """Function to simplify regular expression search"""
    if re.search(re.compile(regular), string):
        return re.search(re.compile(regular, re.I | re.U), string).group()
    else:
        print("There is no any {} pattern in string:\n{}".format(regular, string))


def process_header(header):
    """HEADER EXAMPLE"""
    """
                                                                 13.03.2018
                                                                        ФІТ
Розклад занять на ІІ семестр 2017-2018 н.р.
Напр.(спец.) МІТ
Група ІР
=========================================================
|Тиждень|  ПН   |  ВТ   |  СР   |  ЧТ   |  ПТ   |  СБ   |
=========================================================
| 26.02 |.......|.......|.......|лЛл....|.лЛл...|.......|
| 05.03 |ПППП...|.ЛЛПл..|.ЛПл...|лЛл....|.ЛЛ....|.......|
| 12.03 |ПППП...|.ПЛПл..|лПЛл...|лЛл....|..лл...|.......|
| 19.03 |ПППП...|.ПЛПл..|лПЛл...|лЛл....|.ЛЛ....|.......|
| 26.03 |ПППП...|.ПЛПл..|лПЛл...|лЛл....|..лл...|.......|
| 02.04 |ПППП...|.ПЛП...|лПЛл...|лЛл....|..лл...|.......|
| 09.04 |ПППП...|.ПЛПл..|лПЛл...|лЛл....|лллл...|.......|
| 16.04 |ПППП...|.ПЛПл..|лПЛл...|лЛл....|..лл...|.......|
| 23.04 |ПППП...|.ПЛЛ...|лПЛл...|лЛл....|лллл...|.......|
| 30.04 |ПППП...|.ПЛПл..|лПЛл...|ЛПП....|..лл...|.......|
| 07.05 |ППП....|.ПППл..|.ППл...|ЛПП....|лллл...|.......|
| 14.05 |ППП....|.ПП....|.ПП....|.ПП....|..лл...|.......|
| 21.05 |ППП....|.ППл...|.ПП....|.ПП....|лллл...|.......|
| 28.05 |ППП....|.ППлл..|лПП....|.......|лллл...|.......|
| 04.06 |ППП....|..Плл..|.......|.......|лллл...|.......|
| 11.06 |.......|.......|.......|.......|.......|.......|
| 18.06 |.......|.......|.......|.......|.......|.......|
    """
    group = research(r'(?<=Група )\w{2,3}(-\d{2})*', header)  # (Група ){www-dd}
    year = research(r'\d{4}-\d{4} н.р.', header)  # {dddd-dddd} н.р.
    semester = 1 if research(r'\w+(?= семестр)', header) is 'І' or 'осінній' else 2  # I/II( семестр) new: осінній
    first_monday = research(r'\d{2}.\d{2} ', header)[:-1]  # dd.dd

    # picking first or second number from year variable depending on semester number + day&month from first table cell
    start_date = datetime.date(int(year[0:4]) if semester == 1 else int(year[5:9]),  # year
                               int(first_monday[3:]), int(first_monday[:2]))  # month, day

    print('\nGroup: {}\nYear: {}\nSemester: {}\nFirst week date: {}'
          .format(group, year, semester, start_date))

    return group, year, semester, start_date


def subgroup_number(study_course):
    if re.search(re.compile(r'\d'), study_course):
        subgroup = research(r'\d', study_course)
        study_course = re.sub(re.compile(r'(?<!\d|\()\s*\d\s*(?!\d|\))'), '', study_course)
        study_course = re.sub(re.compile(r'(?<!\d)\s*-\s*(?!\d)'), '', study_course)
        study_course = re.sub(re.compile(r'підгр\.?|підгрупа|група'), '', study_course)
        return subgroup, study_course
    return None, study_course


def room_date(course, year):
    result = []
    for rm in re.split(re.compile('\|'), course)[1:]:
        room = re.findall(re.compile(r'ауд\.\d{3}'), rm)[0]
        room_dates = re.findall(re.compile(r'(?<=\().{5,11}(?=\))'), rm)
        print(room, room_dates)
        for dt in room_dates:
            if len(dt) == 5:
                date = datetime.date(year, int(dt[3:]), int(dt[:2]))
                result.append((room, date))
            else:
                first_date = datetime.date(year, int(dt[3:5]), int(dt[:2]))
                last_date = datetime.date(year, int(dt[9:]), int(dt[6:8]))
                shift = datetime.timedelta(days=7)
                while first_date <= last_date:
                    result.append((room, first_date))
                    first_date += shift
    return result


def process_lessons(schedule, day_desc_list):
    """The function analyzes data about periods and passes it to class Schedule"""
    """LESSONS SECTION EXAMPLE (1st list element of Schedule object)"""
    """
Понеділок
1 пара - 9:00
* Іноземна мова (П) [ас. Бабаніна]
   |ауд.218 (05.03-14.05)|ауд.212 (21.05-04.06)
2 пара - 10:30
* Іноземна мова (П) [ас. Бабаніна]
   |ауд.218 (05.03)|ауд.213 (12.03-30.04)|ауд.218 (07.05)|ауд.213 (14.05)|
   |ауд.205 (21.05)|ауд.204 (28.05-04.06)
3 пара - 12:10
* Іноземна мова (П) [ас. Бабаніна]
   |ауд.204 (05.03-26.03)|ауд.218 (02.04-23.04)|ауд.316 (30.04)|
   |ауд.405 (07.05-14.05)|ауд.218 (21.05)|ауд.204 (28.05-04.06)
4 пара - 13:40
* Іноземна мова (П) [ас. Бабаніна]
   |ауд.204 (05.03-23.04)|ауд.316 (30.04)
    """

    year = schedule.start.year

    print('\nProcessing data...')

    """ Loop through DAYS of week ==================================================================================="""
    for i in range(len(day_desc_list)):
        # Splitting header with "\n" symbol followed by "1 пара"-like pattern
        lessons_of_the_day = re.split(r'\n(?=\d пара)', day_desc_list[i])
        # structure to iter through [day, num+courses_list, num+courses_list, ..., num+courses_list]

        """ Loop through LESSONS ------------------------------------------------------------------------------------"""
        for j in range(1, len(lessons_of_the_day)):
            # Splitting lessons list with "\n* " pattern (* is followed by course name, example: "* Іноземна мова...")
            courses_at_lesson = re.split('\n\* ', lessons_of_the_day[j])
            # structure to iter through [lesson_num, course1, course2, .. ]

            # Information about NUMBER of lesson has INDEX 0!!! (a digit followed by a space and a word with 1+ letter)
            lesson_number = int(research(r'\d(?= \w+)', courses_at_lesson[0]))

            """ Loop through COURSES ********************************************************************************"""
            for crs in range(1, len(courses_at_lesson)):
                # Getting the NAME OF THE COURSE (symbol sequence followed by "(\w)", example: "Іноземна мова (П)")
                study_course = research(r'.+(?=\(\w\))', courses_at_lesson[crs])
                # Excluding speciality from the name of the course (example: "Теорія алгоритмів (МІТ)" => "Теор..тмів ")
                study_course = re.sub(re.compile(r'\(.+\)'), '', study_course)

                subgroup = None
                # If a SUBGROUP number is present, it is assigned to the variable and excluded from the name
                # Examples: "Теорія алгоритмів 1", "Електротехніка та електроніка2", "Технології програмування –2 підгр"
                if re.search(re.compile(r'\d'), study_course):
                    subgroup = research(r'\d', study_course)
                    study_course = re.sub(re.compile(r'\s*\d\s*'), '', study_course)
                    study_course = re.sub(re.compile(r'\s*-\s*'), '', study_course)
                    study_course = re.sub(re.compile(r'підгр\.?|підгрупа|група'), '', study_course)
                # Results example: "Теорія алгоритмів", "Електротехніка та електроніка", "Технології програмування"

                lesson_type = research(r'(?<=\()\w(?=\))', courses_at_lesson[crs])  # (w)
                teacher = research(r'(?<=\[).+(?=\])', courses_at_lesson[crs])  # [name]

                # Getting ROOM numbers and DATES
                rooms_list = re.findall(re.compile(r'ауд.\d{3}( \(.{5,11}\))+'), courses_at_lesson[crs])
                for rm in rooms_list:
                    room = research(r'ауд.\d{3}', rm)

                    room_dates = research(r'(?<=\().{5,11}(?=\))', rm)
                    # Depending on a number of dates 1+ lessons can be added at once
                    if len(room_dates) == 5:
                        dt = datetime.date(year, int(room_dates[3:]), int(room_dates[:2]))
                        schedule.add_lesson(study_course, room, lesson_type, lesson_number, teacher, dt, subgroup)
                    else:
                        list_of_dates = []
                        first_date = datetime.date(year, int(room_dates[3:5]), int(room_dates[:2]))
                        last_date = datetime.date(year, int(room_dates[9:]), int(room_dates[6:8]))
                        shift = datetime.timedelta(days=7)
                        while first_date <= last_date:
                            list_of_dates.append(first_date)
                            first_date += shift
                            for dt in list_of_dates:
                                schedule.add_lesson(study_course, room, lesson_type,
                                                    lesson_number, teacher, dt, subgroup)
    return schedule


def process_data(data):
    """The function processes text from the file and returns a list that is used to create a spreadsheet"""
    # Splitting schedule into sections (0 - general info, 1-5 - days) by "----...----\n" pattern (2+ dashes in line)
    divided_schedule = re.split('-{2,}\n', data)

    schedule_info = process_header(divided_schedule[0])
    scd = Schedule(schedule_info)
    # processing info about classes and passing it to Schedule
    scd = process_lessons(scd, divided_schedule[1:])

    return scd


if __name__ == "__main__":
    lesson1 = """Теорія алгоритмів(л) [проф. Білощицький]
   |ауд.309 (22.05)|ауд.305 (29.05-05.06)"""
    lesson2 = """Теорія систем та системний аналіз (П) [проф. Степанов М.М.]
|ауд.104 (18.10) (01.11) (15.11) (29.11-06.12)"""
    print(lesson2)
    print(room_date(lesson2, 2019))
