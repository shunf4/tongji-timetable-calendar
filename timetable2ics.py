#!/usr/bin/python2
#coding=utf-8
import requests
import urllib3
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)
import sys
import re
import datetime
import time
import getpass
dt = datetime.datetime

cal_Format = '''BEGIN:VCALENDAR
VERSION:2.0
X-WR-CALNAME:同济大学 %(year)s 第 %(semester)s 学期

BEGIN:VTIMEZONE
TZID:Asia/Shanghai
BEGIN:STANDARD
TZOFFSETFROM:+0800
DTSTART:19700101T000000
TZOFFSETTO:+0800
TZNAME:CST
END:STANDARD
END:VTIMEZONE

%(events)s

END:VCALENDAR
'''.decode("utf-8")

event_Format = '''
BEGIN:VEVENT
CLASS:PUBLIC
DESCRIPTION:%(description)s\n
DTSTART;TZID="Asia/Shanghai":%(starttime)s
DTEND;TZID="Asia/Shanghai":%(endtime)s
DTSTAMP:%(stamptime)s
LOCATION:%(location)s
RRULE:FREQ=WEEKLY;COUNT=%(repeatcount)d;INTERVAL=%(interval)d;BYDAY=%(byday)s%(wkstsu)s
SUMMARY:%(coursename)s
UID:%(uid)s

BEGIN:VALARM
ACTION:DISPLAY
DESCRIPTION:Go to classroom
TRIGGER:-P25M
END:VALARM

END:VEVENT
'''.decode("utf-8")

no_Alarm_No_Repeat_Allday_Event_Format = '''
BEGIN:VEVENT
CLASS:PUBLIC
DESCRIPTION:%(description)s\n
DTSTART;VALUE=DATE:%(starttime)s
DTEND;VALUE=DATE:%(endtime)s
DTSTAMP:%(stamptime)s
SUMMARY:%(summary)s
UID:%(uid)s
X-FUNAMBOL-ALLDAY:1
X-MICROSOFT-CDO-ALLDAYEVENT:TRUE
END:VEVENT
'''.decode("utf-8")

datetime_Format = "%Y%m%dT%H%M%S"
onlydate_Format = "%Y%m%d"

timeTable = [
    (datetime.timedelta(), datetime.timedelta()),
    (datetime.timedelta(hours = 8, minutes = 0), datetime.timedelta(hours = 8, minutes = 45)),
    (datetime.timedelta(hours = 8, minutes = 50), datetime.timedelta(hours = 9, minutes = 35)),
    (datetime.timedelta(hours = 10, minutes = 0), datetime.timedelta(hours = 10, minutes = 45)),
    (datetime.timedelta(hours = 10, minutes = 50), datetime.timedelta(hours = 11, minutes = 35)),
    (datetime.timedelta(hours = 13, minutes = 30), datetime.timedelta(hours = 14, minutes = 15)),
    (datetime.timedelta(hours = 14, minutes = 20), datetime.timedelta(hours = 15, minutes = 5)),
    (datetime.timedelta(hours = 15, minutes = 30), datetime.timedelta(hours = 16, minutes = 15)),
    (datetime.timedelta(hours = 16, minutes = 20), datetime.timedelta(hours = 17, minutes = 5)),
    (datetime.timedelta(hours = 17, minutes = 50), datetime.timedelta(hours = 18, minutes = 35)),
    (datetime.timedelta(hours = 19, minutes = 0), datetime.timedelta(hours = 19, minutes = 45)),
    (datetime.timedelta(hours = 19, minutes = 50), datetime.timedelta(hours = 20, minutes = 35)),
    (datetime.timedelta(hours = 20, minutes = 40), datetime.timedelta(hours = 21, minutes = 25)),
]


if sys.stdout.encoding == None:
    reload(sys)
    sys.setdefaultencoding("UTF-8")

def parse_js(expr):
    """
    解析非标准JSON的Javascript字符串，等同于json.loads(JSON str)
    :param expr:非标准JSON的Javascript字符串
    :return:类Python字典
    """
    obj = eval(expr, type('Dummy', (dict,), dict(__getitem__=lambda s, n: n))())
    return obj



proxies = {}
#proxies = { "http": "http://localhost:3408", "https": "https://localhost:3408"}



class Query4m3(object):
    _RE_SAMLResponse = re.compile(r'name="SAMLResponse"\s*value="([\w\d\+\/=]*)"')
    _RE_SAMLURL = re.compile(r';url=([^"]*)"')
    _RE_RelayState = re.compile(r'name="RelayState"\s*value="([\w\d\_\-]*)"')
    _RE_ID = re.compile(r'"ids","(\d+)"')
    _RE_CoursesParse = re.compile(r'<tr>\s*<td>(?P<no>\d+)</td>\s*<td>\s*<a[^>]*>(?P<courseno>[\d\w]+)</a>\s*</td>\s*<td>\s*(?P<coursename>[^<\s]*)\s*(<font[^>]*>[^>]*>)?\s*</td>\s*<td>\s*(?P<compulsory>[^<]*)\s*</td>\s*<td>\s*(?P<examtype>[^<]*)\s*</td>\s*<td>\s*(?P<note>[^<]*)\s*</td>\s*<td>\s*(?P<score>[\d\.]*)\s*</td>\s*<td>\s*(?P<tutor>[^<]*)\s*</td>\s*<td>\s*(?P<timeplace>(?:[^<]|<\s*br\s*>)*)\s*</td>\s*<td>\s*(?P<campus>[^<]*)\s*</td>\s*<td>\s*<a[^>]*>[^>]*</a>\s*</td>\s*</tr>')
    HOST_4M3 = "http://4m3.tongji.edu.cn"
    HOST_IDS = "https://ids.tongji.edu.cn:8443"

    def __init__(self, **kwargs):
        self.w4m3Session = requests.Session()
        self.proxies = kwargs.get("proxies")
        self.w4m3Session.get(self.HOST_4M3, proxies = self.proxies)
        self.headers = {
            "Content-Type": "application/x-www-form-urlencoded; charset=UTF-8", 
            "Connection": "keep-alive",}
        self.id = 0

    def login(self, stuID, passwd):
        # from samlCheck get samlURL to redirect to ids.tongji.edu.cn
        r = self.w4m3Session.get(self.HOST_4M3 + "/eams/samlCheck", proxies = self.proxies)
        samlURL = self._RE_SAMLURL.findall(r.text)[0]

        headers = dict(self.headers).update({"Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,image/apng,*/*;q=0.8","Referer": "http://4m3.tongji.edu.cn/eams/samlCheck"})

        self.w4m3Session.get(samlURL, headers = headers, verify = False, proxies = self.proxies)
        # now w4m3Session has a sessionkey containing info from 4m3.tongji.edu.cn

        # now submit id and pwd
        headers = dict(self.headers).update({"Origin": "https://ids.tongji.edu.cn:8443", "Content-Type": "application/x-www-form-urlencoded", "Referer": "https://ids.tongji.edu.cn:8443/nidp/app/login?id=443&sid=0&option=credential&sid=0"})
        loginURL = self.HOST_IDS + "/nidp/saml2/sso?sid=0&sid=0"
        loginData = {"option": "credential", "Ecom_User_ID": stuID, "Ecom_Password": passwd, "submit": "\xe7\x99\xbb\xe5\xbd\x95"}
        self.w4m3Session.post(loginURL, loginData, headers = headers, verify = False, proxies = self.proxies)

        # jump and get samlvariables
        submitURL = self.HOST_IDS + "/nidp/saml2/sso?sid=0"
        submitData = {}
        headers = dict(self.headers).update({"Host": "ids.tongji.edu.cn:8443", "Referer":  self.HOST_IDS + "/nidp/saml2/sso?sid=0&sid=0"})
        r = self.w4m3Session.post(submitURL, headers = headers, verify = False, proxies = self.proxies)
        
        # now we jump back
        SAMLVarText = r.text

        if len(self._RE_SAMLResponse.findall(SAMLVarText)) == 0:
            raise ValueError, u"Login Failed. ID or password error?"

        SAMLVar = {"SAMLResponse": self._RE_SAMLResponse.findall(SAMLVarText)[0], "RelayState": self._RE_RelayState.findall(SAMLVarText)[0]}
        headers = dict(self.headers).update({"Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,image/apng,*/*;q=0.8"})
        self.w4m3Session.post(self.HOST_4M3 + "/eams/saml/SAMLAssertionConsumer", SAMLVar, headers = headers, proxies = self.proxies)

        # good job! now w4m3session logined.

        headers = dict(self.headers).update({"Referer": "http://4m3.tongji.edu.cn/eams/courseTableForStd!courseTable.action", "Connection": "keep-alive", "X-Requested-With": "XMLHttpRequest", "Host": "4m3.tongji.edu.cn"})
        r = self.w4m3Session.get(self.HOST_4M3 + "/eams/courseTableForStd.action?_=", headers = headers, proxies = self.proxies)

        self.id = self._RE_ID.findall(r.text)[0]

    def getSemesters(self):
        '''
        get semesters.
        return: semesterID, semesterList
        semesterID: current semester id.
        semesterList[0] == ("2007-2008", {1: 47, 2: 48})
        semesterList[0] == (year, {semesterNo: semesterID})
        '''
        semesterQuery = self.HOST_4M3 + "/eams/dataQuery.action"
        semesterData = {"tagId": "semesterBar1485087876Semester", "dataType": "semesterCalendar", "value": "", "empty": False}
        headers = dict(self.headers).update({"Referer": "http://4m3.tongji.edu.cn/", "Origin": "http://4m3.tongji.edu.cn"})
        r = self.w4m3Session.post(semesterQuery, semesterData, headers = headers, proxies = self.proxies)

        #sort dict by year no
        semestersJSON = parse_js(r.text)
        semestersList = list(semestersJSON['semesters'])
        semestersList.sort(lambda y1, y2:cmp(int(y1[1:]), int(y2[1:])))

        resultSemestersList = []
        for i, year in enumerate(semestersList):
            semesterDict = {}
            for semester in semestersJSON['semesters'][year]:
                semesterDict.update({semester['name']: semester['id']})
            resultSemestersList.append((semestersJSON['semesters'][year][0]['schoolYear'].decode("utf-8"), semesterDict))

        return semestersJSON['semesterId'], resultSemestersList

    def getCourseTable(self, semesterID):
        headers = dict(self.headers).update({"Referer": self.HOST_4M3 + "/eams/courseTableForStd!courseTable.action", "X-Requested-With": "XMLHttpRequest"})
        r = self.w4m3Session.post(self.HOST_4M3 + "/eams/courseTableForStd!courseTable.action", {"ignoreHead": 1, "setting.kind": "std", "startWeek": 1, "semester.id": semesterID, "ids": self.id}, headers = headers, proxies = proxies)

        courses = []
        findCourses = self._RE_CoursesParse.finditer(r.text)
        for x in findCourses:
            groupdict = x.groupdict()
            for y in groupdict:
                groupdict[y] = groupdict[y].strip()
            groupdict['timeplace'] = groupdict['timeplace'].split("<br>")
            courses.append(groupdict)

        return courses

def exportICS(loginUsername, loginPassword, nowWeekNo, fileName):
    q = Query4m3(proxies = proxies)
    q.login(loginUsername, loginPassword)
    semesterID, semesterList = q.getSemesters()
    semesterDict = dict(semesterList)
    thisYear = time.localtime().tm_year
    seasonNo = 1
    if (time.localtime().tm_mon < 9):
        seasonNo = 2
        thisYear = thisYear - 1
    termStr = str(thisYear) + "-" + str(thisYear + 1)
    semesterID = semesterDict[termStr][str(seasonNo)]
    courses = q.getCourseTable(semesterID)

    now = dt.now()
    now = now.replace(hour = 0, minute = 0, second = 0, microsecond = 0)
    thisMonday = now + datetime.timedelta(-(now.weekday()))
    firstMonday = thisMonday + datetime.timedelta(-((nowWeekNo - 1) * 7))

    weekdayChar = "一二三四五六日".decode("utf-8")
    weekdayLetter = ["MO", "TU", "WE", "TH", "FR", "SA", "SU"]

    coursesAgendas = []
    for i, course in enumerate(courses):
        try:
            for j, timeplace in enumerate(course['timeplace']):
                infos = timeplace.strip().split(" ")
                if len(infos) > 5:
                    print "Not formatted %s" % ",".join(infos)
                    break

                weekday = weekdayChar.index(infos[1][-1])
                startWeek, endWeek = tuple([int(x) for x in infos[3][infos[3].index("[") + 1:infos[3].index("]")].split('-')])
                startTime = firstMonday + datetime.timedelta(weeks = (startWeek - 1))
                if infos[3][0] != u'[':
                    fortnightOdd = (infos[3][0] == "单".decode("utf-8"))
                    startWeekDelta = fortnightOdd ^ (startWeek % 2)
                    endWeekDelta = fortnightOdd ^ (endWeek % 2)
                    startTime = startTime + datetime.timedelta(weeks = startWeekDelta)
                    repeatTimes = ((endWeek - endWeekDelta) - (startWeek + startWeekDelta)) / 2 + 1
                    interval = 2
                else:
                    repeatTimes = endWeek - startWeek + 1
                    interval = 1
                startTime += datetime.timedelta(days = weekday)
                startPeriod, endPeriod = tuple([int(x) for x in infos[2].split("-")])
                endTime = startTime
                startTime += timeTable[startPeriod][0]
                endTime += timeTable[endPeriod][1]
                
                description_str = "课程编号：%s | 考试考查：%s | 备注：%s | 教师：%s | 学分：%s | 课程性质：%s | 时间地点：%s | 校区：%s".decode("utf-8") % (course['courseno'], course['examtype'], course['note'], course['tutor'], course['score'], course['compulsory'], u"；".join(course['timeplace']), course['campus'])
                startTime_str = startTime.strftime(datetime_Format)
                endTime_str = endTime.strftime(datetime_Format)

                uid_str = "%s-%s-%s@rivage.tk" % (startTime_str, course['courseno'], loginUsername)

                if len(infos) < 5:
                    infos.append("")

                agendaStr = event_Format % {"description": description_str, "starttime": startTime_str, "endtime": endTime_str, "stamptime": dt.now().strftime(datetime_Format), "location": infos[4], "repeatcount": repeatTimes, "interval": interval, "byday": weekdayLetter[weekday], "wkstsu": ";WKST=" + weekdayLetter[0], "coursename": course['coursename'], "uid": uid_str}

                coursesAgendas.append(agendaStr)
        except Exception, e:
            print "Error in %s" % ",".join(course['timeplace'])
            print e

    for i in range(19):
        startTime = firstMonday + datetime.timedelta(weeks = i, hours = 8)
        endTime = firstMonday + datetime.timedelta(weeks = i, hours = 17)
        
        summary_str = "第 %d 周".decode("utf-8") % (i + 1)
        description_str = "第 %d 周开始了！好好加油哦！".decode("utf-8") % (i + 1)
        startTime_str = startTime.strftime(onlydate_Format)
        endTime_str = endTime.strftime(onlydate_Format)

        uid_str = "%s-week%d-%s@rivage.tk" % (startTime_str, i+1, loginUsername)

        agendaStr = no_Alarm_No_Repeat_Allday_Event_Format % {"description": description_str, "starttime": startTime_str, "endtime": endTime_str, "stamptime": dt.now().strftime(datetime_Format), "summary": summary_str, "uid": uid_str}

        coursesAgendas.append(agendaStr)

    iCSStr = cal_Format % {"year": termStr, "semester": seasonNo, "events": u'\n'.join(coursesAgendas)}

    f = open(fileName, "w")
    f.write(iCSStr.encode("utf-8"))
    f.close()

if __name__ == "__main__":
    loginUsername = raw_input("Unified Identification(ID):")
    loginPassword = getpass.getpass("Password(Invisible input):")
    nowWeekNo = int(raw_input("Week Number now(Sunday is the last day of a week):"))
    fileName = "./%s.ics" % loginUsername
    exportICS(loginUsername, loginPassword, nowWeekNo, fileName)
    print("Generated " + fileName + ".")

