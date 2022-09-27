import requests
import datetime


try:
    test = requests.request(
        method='POST',
        url='https://querycourse.ntust.edu.tw/querycourse/api/courses'
    )

    test.raise_for_status()

except requests.exceptions.HTTPError as e:

    print("\nError Occurred: Please make sure that NTUST Course is still functioning!")
    raise SystemExit(e)

else:
    print('NTUST Connection Test Pass!')

print('\nWelcome to NTUST Course to iCalender Transfer Tool!\n')

while True:

    Start = input('Enter Semester Start Date (ex. 2022.9.5) >> ')
    ss = Start.split('.')
    if len(ss) != 3 or Start[-1:] == '.':
        print('Wrong Answer!\n')
        continue
    elif len(str(int(ss[0]))) != 4 and len(str(int(ss[1]))) != 2 and len(str(int(ss[2]))) != 2:
        print('Wrong Digit!\n')
        continue
    StartDate = datetime.datetime.strptime(Start, '%Y.%m.%d')

    End = input('Enter Semester End Date (ex. 2022.12.23) >> ')
    ee = End.split('.')
    if len(ee) != 3 or End[-1:] == '.':
        print('Wrong Answer!\n')
        continue
    elif len(str(int(ee[0]))) != 4 and len(str(int(ee[1]))) != 2 and len(str(int(ee[2]))) != 2:
        print('Wrong Digit!\n')
        continue
    EndDate = datetime.datetime.strptime(End, '%Y.%m.%d')

    if EndDate <= StartDate or EndDate < datetime.datetime.now():
        print('Entry Not Valid! Please Re-Enter')
    else:
        break

if int(StartDate.strftime('%m')) >= 6:
    # 上學期
    SemCode = str(int(StartDate.strftime('%Y')) - 1911) + '1'
else:
    # 下學期
    SemCode = str(int(StartDate.strftime('%Y')) - 1912) + '2'

print('Semester might be ' + SemCode)

# print('This is', t.strftime('%B %Y'), ', By Default, Semester Code Should be', SemCode)

'''i = input('Press "ENTER" to Select Semester ' + SemCode + ' or key in to change >> ')
if i != '':
    while len(str(int(i))) != 4:
        i = input('Please type the correct Semester Code (e.g. 1101) >> ')

    while i[3] != '1' and i[3] != '2':
        i = input('Please type the correct Semester Code (e.g. 1101) >> ')

    SemCode = i
    print('Successfully change Semester Code to ', SemCode)'''

i = ''
courses = []
print('\n ~~ ENTER COURSE ID ~~ \nStart with "-" for removal.(e.g. -EE12345678) \nCommand "list" for list all selected. '
      '\nCommand "clear" to clean selected. \nWhen Finish, Key in "end" to next step.')
while i != 'end':
    i = input('>> ')
    i = i.strip(' \n\t')  # Remove Weird Spaces

    # Check All Special Functions
    if '-' in i:
        flag = False
        i = i.replace('-', "")
        for _ in courses:
            if _[0]['CourseNo'] == i:
                print('Removed', _[0]['CourseName'])
                courses.remove(_)
                flag = True

        if not flag:
            print('Not Found!')

    elif i == '':
        pass

    elif i == 'end':
        print('Following Courses Are Selected. Is it Correct?')
        for _ in courses:
            print(_[0]['CourseNo'] + ' ' + _[0]['CourseName'])
        c = input('\nPress "ENTER" to continue. Key "n" to continue entry. >> ')
        if c == 'n':
            i = ''
            pass

    elif i == 'list':
        if not courses:
            print('None')
        else:
            for _ in courses:
                print(_[0]['CourseNo'] + ' ' + _[0]['CourseName'])

    elif i == 'clear':
        print('All Clean!')
        courses = []

    else:
        r = requests.post(
            url='https://querycourse.ntust.edu.tw/querycourse/api/courses',
            data={
                "Semester": SemCode,
                "CourseNo": i,
                "Language": "zh"
            }
        )

        data = r.json()
        '''
        中文課名搜尋 還不想處理
        if not data:
            r = requests.request(
                method='POST',
                url='https://querycourse.ntust.edu.tw/querycourse/api/courses',
                data={
                    "Semester": SemCode,
                    "CourseName": i,
                    "Language": "zh"
                }
            )
            data = r.json()'''

        if not data:
            print("Course Not Found!")
        elif data in courses:
            print('Already Exists!')
        else:
            print("Selected ", data[0]['CourseName'], 'on', data[0]['Node'])
            courses.append(data)

    ### Parse NTUST Week Code ###
WeekPair = {
    'M': '1', 'T': '2', 'W': '3', 'R': '4', 'F': '5', 'S': '6', 'U': '7'
}
StartTimePair = {
    '1': '081000', '2': '091000', '3': '102000', '4': '112000', '5': '122000',
    '6': '132000', '7': '142000', '8': '153000', '9': '163000', '10': '173000',
    'A': '182500', 'B': '192000', 'C': '201500', 'D': '211000'
}
EndTimePair = {
    '1': '090000', '2': '100000', '3': '111000', '4': '121000', '5': '131000',
    '6': '141000', '7': '151000', '8': '162000', '9': '172000', '10': '182000',
    'A': '191500', 'B': '201000', 'C': '210500', 'D': '220000'
}
WeekTime = []
for course in courses:
    course[0]['WeekTime'] = course[0]['Node'].split(',')
    course[0]['Day'] = []
    course[0]['StartTime'] = []
    course[0]['EndTime'] = []
    course[0]['Index'] = 0
    #course[0].pop('Node')
    for WT in course[0]['WeekTime']:

        course[0]['Day'].append(WeekPair[WT[0]])
        course[0]['StartTime'].append(StartTimePair[WT[1:]])
        course[0]['EndTime'].append(EndTimePair[WT[1:]])
        course[0]['Index'] += 1

courses.sort(key=lambda x: x[0]['Day'])

try:
    print('Creating File...')
    f = open('output.ics', 'w', encoding="UTF-8")
except IOError:
    print('File Failed')
else:
    print('Writing File...')
    f.write('BEGIN:VCALENDAR\n')
    f.write('PRODID:JianYinChian\n')
    f.write('VERSION:2.0\n')
    f.write('X-WR-CALNAME:Course' + SemCode + '\n')
    f.write('X-WR-TIMEZONE:Asia/Taipei\n')
    f.write('BEGIN:VTIMEZONE\n')
    f.write('TZID:Asia/Taipei\n')
    f.write('X-LIC-LOCATION:Asia/Taipei\n')
    f.write('BEGIN:STANDARD\n')
    f.write('TZOFFSETFROM:+0800\n')
    f.write('TZOFFSETTO:+0800\n')
    f.write('TZNAME:CST\n')
    f.write('DTSTART:19700101T000000\n')
    f.write('END:STANDARD\n')
    f.write('END:VTIMEZONE\n')
    delta = datetime.timedelta(days=1)
    Date = StartDate
    i = 0
    while Date <= EndDate:
        #print(Date.strftime("%Y%m%d"))
        for course in courses:
            for _ in range(course[0]['Index']):
                #print(course[0]['Day'][_], Date.strftime('%w'))
                if course[0]['Day'][_] == Date.strftime('%w'):
                    print('Adding ' + course[0]['CourseNo'] + ' on ' + Date.strftime('%A'))
                    i += 1
                    Now = datetime.datetime.now()
                    f.write('BEGIN:VEVENT\n')
                    f.write('SUMMARY:' + course[0]['CourseName'] + '\n')
                    f.write('DESCRIPTION:' + course[0]['CourseTeacher'] + '\n')
                    f.write('DTSTART;TZID=Asia/Taipei:' + Date.strftime("%Y%m%d") + 'T' + course[0]['StartTime'][_] + '\n')
                    f.write('DTEND;TZID=Asia/Taipei:' + Date.strftime("%Y%m%d") + 'T' + course[0]['EndTime'][_] + '\n')
                    f.write('LOCATION:' + course[0]['ClassRoomNo'] + '\n')

                    f.write('DTSTAMP:' + Now.strftime("%Y%m%d") + 'T' + Now.strftime("%H%M%S") + '\n')
                    f.write('UID:IDontKnowWhatIsThis' + str(i) + '\n')

                    f.write('END:VEVENT\n')

        Date += delta
    f.write('END:VCALENDAR')
    f.close()
    print('\nDone! Please Check "Output.ics" and import file to Calender APP!')


