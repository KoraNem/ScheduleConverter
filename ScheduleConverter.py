"""
The program processes schedule, which is given in the text format, and converts it into spreadsheet (.xlsx file)
"""

import re
import datetime
import openpyxl
from openpyxl.styles import Border, Side, Alignment, Font, Color, NamedStyle, PatternFill


"""
sch.weeks = [week = [day.periods=[period,period,period,period,...], day,day,day,day], [],[],[],[],...]
"""


class Schedule:
    update = None
    year = None
    group = None
    semester = None
    start = None
    weeks = [[[None for k in range(7)] for j in range(5)] for i in range(17)]

    def __init__(self, info):
        # adding attributes
        self.group = info[0]
        self.year = info[1]
        self.semester = info[2]
        self.start = info[3]

    def add_period(self, name, room, p_type, number, teacher, date, sub):
        # calculating week and day indexes to access the necessary cell in the weeks attribute
        num_of_w = int((date - self.start).days) // 7
        day_of_w = int((date - self.start).days) % 7
        self.weeks[num_of_w][day_of_w][number - 1] = Class(name, room, p_type, number, teacher, date, sub)
        return self

    def show(self):
        """
        функція відображає вміст поля weeks для перевірки коректності роботи програми
        """
        print('Group: ', self.group)
        print('Year: ', self.year)
        print('Semester: ', self.semester)
        print('First day date: ', self.start)
        for i in range(17):  # weeks in attributes
            print('\nweek {}'.format(i + 1))
            for j in range(5):  # days in week
                print('day {}'.format(j + 1))
                print(self.weeks[i][j])

    def export(self):
        """
        функція готує дані до передачі в файл xlsx
        формат:
            назва пари
            аудиторія, тип, група
            назва
            аудиторія
            ...

        цикл по 1х парах 5ти днів
        цикл по аудиторіях 5ти днів
        """
        export_data = []
        for i in range(17):
            for j in range(7):
                row_name = []
                row_info = []
                for k in range(5):
                    if self.weeks[i][k][j]:
                        row_name.append(self.weeks[i][k][j].discipline)
                        if self.weeks[i][k][j].subgroup:
                            row_info.append(self.weeks[i][k][j].room + ', група '
                                            + self.weeks[i][k][j].subgroup)
                        else:
                            row_info.append(self.weeks[i][k][j].room)
                    else:
                        row_name.append('')
                        row_info.append('')
                export_data.append(row_name)
                export_data.append(row_info)
        return export_data


class Class:
    discipline = None   # назва дисципліни
    room = None         # номер пари
    type = None         # тип пари: Л - лекція; П - практика; л - лабораторна.
    number = None       # порядковий номер пари
    teacher = None      # ім'я викладача
    subgroup = None     # підгрупа
    date = None         # дата проведення пари

    def __init__(self, discipline, room, c_type, number, teacher, date, subgroup=None):
        self.discipline = discipline
        self.room = room
        self.type = c_type
        self.number = number
        self.teacher = teacher
        self.subgroup = subgroup
        self.date = date


def import_from_file(directory):
    """
    function reads file and assigns its contents to variable schedule
    """
    txt = open(directory, "r", encoding='windows-1251')
    schedule = txt.read()
    txt.close()
    return schedule


def process_data(data):
    """
    this function processes given data and returns it in a suitable for xlsx format form
    it uses Schedule object to store the data and return the list that will be returned to spreadsheet
    """

    def get_header_info(header):
        """
        function extracts data from header using regex module
        return group, year, semester, start_date
        """
        group = research(r'(?<=Група )\w{2,3}(-\d{2})*', header)
        year = research(r'\d{4}-\d{4} н.р.', header)
        semester = 1 if research(r'\w+(?= семестр)', header) is 'І' else 2
        first_monday = research(r'\d{2}.\d{2} ', header)[:-1]
        start_date = datetime.date(int(year[0:4]) if semester == 1 else int(year[5:9]),
                                   int(first_monday[3:]), int(first_monday[:2]))
        print('\nGroup: {}\nYear: {}\nSemester: {}\nFirst week date: {}'
              .format(group, year, semester, start_date))
        return group, year, semester, start_date

    def get_classes(sc, list_of_days):
        """function analyzes data about periods and passes it to class Schedule"""
        # days loop
        print('\nProcessing data...')
        for i in range(len(list_of_days)):
            day = re.split(r'\n(?=\d пара)', list_of_days[i])  # day = [day, num+disclist, num+disclist, ...]
            # one day (classes of the day) loop
            for j in range(1, len(day)):
                dis = re.split('\n\* ', day[j])  # dis = [num, disc1, disc2, .. ]
                # information about NUMBER of class has INDEX 0 in list dis!!!
                for k in range(1, len(dis)):
                    # discipline, room, type, number, teach, date
                    number = int(research(r'\d(?= \w+)', dis[0]))

                    # there is a possibility that the name contains group number, so that we have to analyze the name
                    name = research(r'.+(?= \(\w\))', dis[k])
                    name = re.sub(re.compile(r'\(.+\)'), '', name)
                    subgroup = None
                    if re.search(re.compile(r'\d'), name):
                        subgroup = research(r'\d', name)
                        name = re.sub(re.compile(r'\s*\d\s*'), '', name)
                        name = re.sub(re.compile(r'підгр?|підгрупа|група'), '', name)

                    c_type = research(r'(?<=\()\w(?=\))', dis[k])
                    teacher = research(r'(?<=\[).+(?=\])', dis[k])
                    rooms = re.findall(re.compile(r'ауд.\d{3} \(.{5,11}\)'), dis[k])
                    for n in rooms:
                        room = research(r'ауд.\d{3}', n)
                        date_s = research(r'(?<=\().{5,11}(?=\))', n)
                        if len(date_s) == 5:
                            d = datetime.date(2018, int(date_s[3:]), int(date_s[:2]))
                            sc.add_period(name, room, c_type, number, teacher, d, subgroup)
                        else:
                            dates = []
                            s_date = datetime.date(2018, int(date_s[3:5]), int(date_s[:2]))
                            l_date = datetime.date(2018, int(date_s[9:]), int(date_s[6:8]))
                            shift = datetime.timedelta(days=7)
                            while s_date <= l_date:
                                dates.append(s_date)
                                s_date += shift
                                for t in dates:
                                    sc.add_period(name, room, c_type, number, teacher, t, subgroup)
        return sc

    def research(regular, string):
        return re.search(re.compile(regular), string).group()

    # splitting schedule into sections (0 - general info, 1-5 - days)
    divided_sched = re.split('-{2,}\n', data)
    # processing data from the first section
    schedule_info = get_header_info(divided_sched[0])
    # creating Schedule object by passing it data from the firs section
    scd = Schedule(schedule_info)
    # processing info about classes and passing it to Schedule
    scd = get_classes(scd, divided_sched[1:])
    # returning list of classes + info about schedule
    exp = scd.export()
    del scd
    return exp, schedule_info


def export_to_excel(ex_data):
    """function takes ready data, creates xlsx file and passes all the information to it"""

    def workbook_styles():
        # COLORS
        blk = Color('000000')
        ttl = Color('992600')
        col_stls = (Color('ffffcc'), Color('e6ffcc'))

        # ALIGNMENT
        al_center = Alignment(horizontal='center', vertical='center', wrap_text=True)
        al_cent_rot = Alignment(horizontal='center', vertical='center', wrap_text=False, text_rotation=90)
        al_bottom = Alignment(horizontal='center', vertical='bottom', wrap_text=True)
        al_top = Alignment(horizontal='center', vertical='top', wrap_text=True)

        # BORDERS
        side = Side(color=blk, style='thin')
        all_borders = Border(top=side, bottom=side, left=side, right=side)
        without_bot = Border(top=side, left=side, right=side)
        without_top = Border(bottom=side, left=side, right=side)

        # FONTS
        f_title = Font(name='Calibri Light', size=12, bold=True, color=ttl)
        f_subtitle = Font(name='Calibri Light', size=10, color=ttl)
        f_day_date = Font(name='Calibri Light', size=10)
        f_disc = Font(name='Calibri', size=11)
        f_room = Font(name='Calibri', size=9)

        # STYLES
        s_title = NamedStyle(name='title', font=f_title, alignment=al_bottom, border=without_bot)
        s_sub = NamedStyle(name='subtitle', font=f_subtitle, alignment=al_top, border=without_top)
        s_day = NamedStyle(name='day', font=f_day_date, alignment=al_center, border=all_borders)
        s_date = NamedStyle(name='date', font=f_day_date, alignment=al_cent_rot, border=all_borders)
        s_disc = NamedStyle(name='disc', font=f_disc, alignment=al_center, border=without_bot)
        s_room = NamedStyle(name='room', font=f_room, alignment=al_top, border=without_top)
        s_num = NamedStyle(name='snum', font=f_day_date, alignment=al_center, border=all_borders)

        return s_title, s_sub, s_day, s_date, s_disc, s_room, s_num, col_stls

    def style(sheet, start, i, disc, room, num, date, fill1, fill2, f):
        sheet.merge_cells('A{}:A{}'.format(start, i + 1))
        for j in range(start, i + 2, 2):
            # classes cells
            for k in sheet[j]:
                k.style = disc
                k.fill = fill1 if f % 2 == 0 else fill2
            for k in sheet[j + 1]:
                k.style = room
                k.fill = fill1 if f % 2 == 0 else fill2

            # date cells styling
            sheet['A' + str(j)].style = date
            sheet['A' + str(j)].number_format = 'DD.MM'
            sheet['A' + str(j + 1)].style = date

            # number cells styling
            sheet.merge_cells('B{}:B{}'.format(j, j + 1))
            sheet['B' + str(j)].style = num
            sheet['B' + str(j + 1)].style = num

        return sheet

    # IMPORTING DATA
    schedule_properties = ex_data[1]  # group, year, semester, start
    classes = ex_data[0]

    # CREATING WORKBOOK
    print('Creating spreadsheet...')
    filename = schedule_properties[0] + '.xlsx'
    s_wb = openpyxl.workbook.Workbook()
    sheet = s_wb.active
    sheet.title = schedule_properties[0]

    # ADDING INFORMATION
    # classes
    days = ['Понеділок', 'Вівторок', 'Середа', 'Четвер', 'П\'ятниця']
    sheet.append(days)
    for row in classes:
        sheet.append(row)
    # date + number
    sheet.insert_cols(0, 2)
    curr_date = schedule_properties[3]
    shift = datetime.timedelta(days=7)
    i = 2
    while sheet['C' + str(i)].value is not None:
        for j in range(0, 14, 2):
            sheet['A' + str(i + j)] = curr_date
            sheet['B' + str(i + j)] = (j + 2) // 2
        curr_date += shift
        i += 14
    # header
    sheet.insert_rows(0, 2)
    sheet['A1'] = 'РОЗКЛАД НА {} СЕМЕСТР'.format(schedule_properties[2])
    sheet['A2'] = 'Група {}, {}'.format(schedule_properties[0], schedule_properties[1])

    # IMPORTING STYLES
    title, subtitle, day, date, disc, room, num, col_st = workbook_styles()
    s_wb.add_named_style(title)
    s_wb.add_named_style(subtitle)
    s_wb.add_named_style(day)
    s_wb.add_named_style(date)
    s_wb.add_named_style(disc)
    s_wb.add_named_style(room)
    s_wb.add_named_style(num)
    col1, col2 = col_st
    # fills
    fill1 = PatternFill(fgColor=col1, fill_type='solid')
    fill2 = PatternFill(fgColor=col2, fill_type='solid')

    print(sheet)

    # DELETING EMPTY ROWS, APPLYING STYLES AND MERGING CELLS
    # dimensions
    for i in range(2):
        sheet.column_dimensions[chr(ord('A') + i)].width = 4
    for i in range(5):
        sheet.column_dimensions[chr(ord('C') + i)].width = 19

    # title
    sheet.merge_cells('A1:G1')
    sheet.merge_cells('A2:G2')
    sheet.merge_cells('A3:B3')
    for i in sheet[1]:
        i.style = title
    for i in sheet[2]:
        i.style = subtitle
    for i in sheet[3]:
        i.style = day

    # classes
    # deleting rows
    i = 4
    while sheet['A' + str(i)].value is not None:
        if (sheet['C' + str(i)].value is ''
                and sheet['D' + str(i)].value is ''
                and sheet['E' + str(i)].value is ''
                and sheet['F' + str(i)].value is ''
                and sheet['G' + str(i)].value is ''):
            sheet.delete_rows(i, 2)
        else:
            i += 2

    # applying styles
    i = 4
    start = 4
    f = 0
    while sheet['B' + str(i)].value is not None:
        if sheet['B' + str(i + 2)].value is None:
            sheet = style(sheet, start, i, disc, room, num, date, fill1, fill2, f)
        elif sheet['B' + str(i + 2)].value < sheet['B' + str(i)].value:
            sheet = style(sheet, start, i, disc, room, num, date, fill1, fill2, f)
            start = i + 2
            f += 1
        i += 2

    print('File {} is successfully saved'.format(filename))
    s_wb.save(filename)
    del s_wb


"""
Program
"""

while True:
    dir = input('Enter the file name (path): ')
    try:
        dat = import_from_file(dir)
    except FileNotFoundError:
        print('Oops! It seems like there is no such file.')
        ans = input('Do you want to try again? (Y)es or (N)o? ')
        if ans != 'Y' or ans != 'y':
            break
        print('\n')
        continue
    n = process_data(dat)
    export_to_excel(n)
    break
