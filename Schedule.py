class Lesson:
    def __init__(self, course, room, l_type, number, teacher, date, subgroup=None):
        self.course = course  # study course | назва дисципліни
        self.room = room  # room number | номер аудиторії
        self.l_type = l_type  # lesson type | тип пари: Л - лекція; П - практика; л - лабораторна.
        self.number = number  # lesson sequence number | порядковий номер пари
        self.teacher = teacher  # teacher | ім'я викладача
        self.date = date  # lesson date | дата проведення пари
        self.subgroup = subgroup  # subgroup number| підгрупа


class Schedule:
    """
    This class holds the schedule structure and its description
    """
    group = None
    year = None
    semester = None
    start = None  # date of the first lesson
    lessons_list = [[[None for lesson in range(7)] for day in range(5)] for week in range(20)]

    def __init__(self, info):
        # main attributes of schedule
        self.group = info[0]
        self.year = info[1]
        self.semester = info[2]
        self.start = info[3]

    @staticmethod
    def get_lesson_type(ls_type):
        if ls_type == 'Л' or ls_type == 'L':
            return 'лекція'
        elif ls_type.lower() == 'п' or ls_type.lower() == 'p':
            return 'практика'
        elif ls_type.lower() == 'с' or ls_type.lower() == 'c':
            return 'семінар'
        elif ls_type == 'л' or ls_type == 'l':
            return 'лабораторна'
        else:
            return ls_type

    def add_lesson(self, study_course, room, lesson_type, lesson_number, teacher, date, subgroup):
        # calculating week and day indexes to access the necessary cell (weeks attribute)
        index_week = int((date - self.start).days) // 7
        index_day = int((date - self.start).days) % 7
        current_lesson = Lesson(study_course, room, lesson_type, lesson_number, teacher, date, subgroup)
        self.lessons_list[index_week][index_day][lesson_number - 1] = current_lesson
        return self

    def show(self):
        """функція відображає вміст поля weeks для перевірки коректності роботи програми"""
        print('Group: ', self.group)
        print('Year: ', self.year)
        print('Semester: ', self.semester)
        print('First day date: ', self.start)
        for i in range(17):  # weeks in attributes
            print('\nweek {}'.format(i + 1))
            for j in range(5):  # days in week
                print('day {}'.format(j + 1))
                print(self.lessons_list[i][j])

    def create_spreadsheet(self):
        """ This method returns a list that will be used to create a spreadsheet
            The result table is formed row by row. Every two rows correspond a single week.

                * THE STRUCTURE OF A LIST RETURNED BY THE FUNCTION *
                +==============+==============+==============+==============+==============+
                |    course    |    course    |    course    |    course    |    course    |    rowCourseName
                +--------------+--------------+--------------+--------------+--------------+
                |room, subgroup|room, subgroup|room, subgroup|room, subgroup|room, subgroup|    rowLessonInfo
                +==============+==============+==============+==============+==============+
                |    course    |    course    |    course    |    course    |    course    |    rowCourseName
                +--------------+--------------+--------------+--------------+--------------+
                |room, subgroup|room, subgroup|room, subgroup|room, subgroup|room, subgroup|    rowLessonInfo
                +==============+==============+==============+==============+==============+

                * THIS IS HOW IT WILL LOOK LIKE AFTER WEEK DATES ARE ADDED *
                week 0
                +====================================+                      цикл по тижнях:
                |===========+========================+                         цикл по парах:
                |           |day0|day1|day2|day3|day4| rowCourseName                цикл по днях:
                |  lesson 0 |----+----+----+----+----+                                  1. назва пари
                |           |day0|day1|day2|day3|day4| rowLessonInfo                    2. аудиторія + підгрупа
                |===========+========================+
                |           |day0|day1|day2|day3|day4| rowCourseName
                |  lesson 1 |----+----+----+----+----+
                |           |day0|day1|day2|day3|day4| rowLessonInfo
                |===========+========================+
                +====================================+
        """
        spreadsheet_lessons = []

        for week in range(17):
            for lesson in range(7):
                row_course = []
                row_lesson_info = []

                for day in range(5):
                    if self.lessons_list[week][day][lesson]:
                        # lesson name
                        row_course.append(self.lessons_list[week][day][lesson].course)
                        # lesson information (room+type+subgroup(optional))
                        temp_info = (self.lessons_list[week][day][lesson].room + ', '
                                     + self.lessons_list[week][day][lesson].l_type)
                        if self.lessons_list[week][day][lesson].subgroup:
                            temp_info += ', підгрупа ' + self.lessons_list[week][day][lesson].subgroup
                        row_lesson_info.append(temp_info)
                    else:
                        row_course.append('')
                        row_lesson_info.append('')

                spreadsheet_lessons.append(row_course)
                spreadsheet_lessons.append(row_lesson_info)

        return spreadsheet_lessons
