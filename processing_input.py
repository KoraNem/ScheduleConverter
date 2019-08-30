import datetime
import re
from Schedule import Schedule


def research(regular, string):
    return re.search(re.compile(regular), string).group()


def get_header_info(header):
    group = research(r'(?<=Група )\w{2,3}(-\d{2})*', header)                    # (Група ){www-dd}
    year = research(r'\d{4}-\d{4} н.р.', header)                                # {dddd-dddd} н.р.
    semester = 1 if research(r'\w+(?= семестр)', header) is 'І' else 2          # I/II( семестр)
    first_monday = research(r'\d{2}.\d{2} ', header)[:-1]                       # dd.dd\

    start_date = datetime.date(int(year[0:4]) if semester == 1
                               else int(year[5:9]), int(first_monday[3:]), int(first_monday[:2]))

    print('\nGroup: {}\nYear: {}\nSemester: {}\nFirst week date: {}'
          .format(group, year, semester, start_date))

    return group, year, semester, start_date


def get_classes(schedule, day_desc_list):
    """The function analyzes data about periods and passes it to class Schedule"""
    year = schedule.start.year

    print('\nProcessing data...')

    # Loop through DAYS of week ====================================================================================
    for i in range(len(day_desc_list)):
        lessons_of_the_day = re.split(r'\n(?=\d пара)', day_desc_list[i])  # [day, num+disclist, num+disclist, ...]

        # Loop through LESSONS -------------------------------------------------------------------------------------
        for j in range(1, len(lessons_of_the_day)):
            courses_at_lesson = re.split('\n\* ', lessons_of_the_day[j])  # [num, disc1, disc2, .. ]
            # information about NUMBER of lesson has INDEX 0!!!
            lesson_number = int(research(r'\d(?= \w+)', courses_at_lesson[0]))

            # Loop through COURSES *********************************************************************************
            for crs in range(1, len(courses_at_lesson)):
                # Getting the name of the COURSE
                study_course = research(r'.+(?= \(\w\))', courses_at_lesson[crs])
                # Excluding speciality from name of the course
                study_course = re.sub(re.compile(r'\(.+\)'), '', study_course)
                subgroup = None
                # If a subgroup number is present, it is assigned to the variable and excluded from the name
                if re.search(re.compile(r'\d'), study_course):
                    subgroup = research(r'\d', study_course)
                    study_course = re.sub(re.compile(r'\s*\d\s*'), '', study_course)
                    study_course = re.sub(re.compile(r'підгр?|підгрупа|група'), '', study_course)

                lesson_type = research(r'(?<=\()\w(?=\))', courses_at_lesson[crs]) # (w)
                teacher = research(r'(?<=\[).+(?=\])', courses_at_lesson[crs]) # [name]

                # Getting room numbers and dates
                rooms_list = re.findall(re.compile(r'ауд.\d{3} \(.{5,11}\)'), courses_at_lesson[crs])
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
    # Splitting schedule into sections (0 - general info, 1-5 - days)
    divided_schedule = re.split('-{2,}\n', data)

    schedule_info = get_header_info(divided_schedule[0])
    scd = Schedule(schedule_info)
    # processing info about classes and passing it to Schedule
    scd = get_classes(scd, divided_schedule[1:])

    return scd
