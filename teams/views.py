from datetime import datetime, date
from django.db.models import Q
from django.http import HttpResponse
from django.shortcuts import render, redirect
from teams.Modules.GenerateReport import GenerateReport
from teams.Modules.GetReport import GetReport
from teams.models import Login, Team, workon, DailyAttendance, ShiftTimings
import logging
logger = logging.getLogger('django')

# Create your views here.
def login(request):

    usernames = Login.objects.all()
    for username in usernames:
        print(username.username)
    context = {'usernames':usernames}
    return render(request, 'login.html', context=context)

def logout(request):
    del request.session['username']
    del request.session['email']
    usernames = Login.objects.all()
    for username in usernames:
        print(username.username)
    context = {'usernames': usernames}
    return render(request, 'login.html', context=context)

def team(request):
    team = Team.objects.filter(Q(Tlmail=request.session['email']))
    Context ={'team':team}
    return render(request, 'team.html',context=Context)

def updateattendance(request):
    attendancedate = request.POST["attendancedate"]
    employeeno = request.POST["employeeno"]
    employeename = request.POST["employeename"]
    worktype = request.POST["workon"]
    shift = request.POST["shift"]
    teamcount = Team.objects.filter(Tlmail=request.session['email']).count()

    input_date = datetime.strptime(attendancedate, "%d-%m-%Y")
    DailyAttendance.objects.filter(attendancedate__date=input_date, Employeeno=employeeno).update(
        attendancedate=input_date,
        worktype=worktype,
        Shift = shift
    )

    team = DailyAttendance.objects.filter(Q(Tlmail=request.session['email']))
    for member in team:
        teamname = member.Teamname
        input_datetime = datetime.strptime(str(member.attendancedate), "%Y-%m-%d %H:%M:%S%z")
        member.attendancedate = input_datetime.strftime("%d-%m-%Y")

    message="DataUpdated"
    Context = {'teamcount':teamcount,'message':message,'team': team, 'username': request.session['username'], 'teamname': teamname}
    return render(request, 'Read.html', context=Context)

def editattendance(request):
    team = DailyAttendance.objects.filter(Q(Tlmail=request.session['email']))
    teamcount = Team.objects.filter(Tlmail=request.session['email']).count()
    for member in team:
        teamname = member.Teamname
    editdetails =  str(request.POST["editdetail"])
    editdate = editdetails.split("/")[0]
    editempno = editdetails.split("/")[1]
    input_date = datetime.strptime(editdate , "%d-%m-%Y")
    record = DailyAttendance.objects.filter(attendancedate=input_date, Employeeno=editempno)
    work = workon.objects.all()

    for employee in record:
        input_datetime = datetime.strptime(str(employee.attendancedate), "%Y-%m-%d %H:%M:%S%z")
        employee.attendancedate = input_datetime.strftime("%d-%m-%Y")

    shiftimings = ShiftTimings.objects.all()
    Context = {'teamcount':teamcount,'username': request.session['username'],'shiftimings':shiftimings,'record':record, 'teamname': teamname,'work':work}
    return render(request,'editattendance.html',context=Context)

def readdata(request):
    team = DailyAttendance.objects.filter(Q(Tlmail=request.session['email'])).order_by('-attendancedate')[:500]

    teamcount = Team.objects.filter(Tlmail=request.session['email']).count()

    for member in team:
        teamname = member.Teamname
        input_datetime = datetime.strptime(str(member.attendancedate), "%Y-%m-%d %H:%M:%S%z")
        member.attendancedate = input_datetime.strftime("%d-%m-%Y")
        print(member.attendancedate)

    if team.exists():
        Context = { 'teamcount':teamcount,'team': team, 'username': request.session['username'], 'teamname': teamname}
    else:
        Context = {'norecords':'nodata'}

    return render(request, 'Read.html', context=Context)

def processlogin(request):

    email = request.POST["Email"]
    request.session['email']=email
    password = request.POST["Password"]
    login_query = Login.objects.all()
    shiftimings = ShiftTimings.objects.all()
    loginbool = False

    logger.info(str(email)+str(" loggedin ")+ str(datetime.now()))

    for login in login_query:
        if str(login.username).upper()==str(email).upper() and str(login.password).upper()==str(password).upper():
            loginbool = True

    if loginbool == True and (str(email).upper()=="SATYANARAYANA.P@REDINGTONGROUP.COM" or str(email).upper()=="DEEPAK.KUMARU@REDINGTONGROUP.COM" or str(email).upper()=="UDHAYASHANKAR.S@REDINGTONGROUP.COM" ):
        request.session['username'] = str(email).split("@")[0].split(".")[0]
        Context = {'username': request.session['username']}
        return render(request, 'Report.html', context=Context)
    elif loginbool == True:
        request.session['username'] = str(email).split("@")[0].split(".")[0]
        team =  Team.objects.filter(Q(Tlmail=email))
        teamcount = Team.objects.filter(Tlmail=email).count()
        work = workon.objects.all()
        for member in team:
            teamname = member.Teamname

        if team.exists():
            Context = {'teamcount':teamcount,'shiftimings': shiftimings, 'username': request.session['username'], 'team': team, 'work': work,'teamname': teamname}
        else:
            Context = {'norecords':'nodata'}

        return render(request, 'index.html', context=Context)
    else:
        usernames = Login.objects.all()
        for username in usernames:
            print(username.username)
        context = {'usernames': usernames}
        return render(request, 'login.html',context=context)


def home(request):
    work = workon.objects.all()
    team = Team.objects.filter(Q(Tlmail=request.session['email']))
    teamcount = Team.objects.filter(Tlmail=request.session['email']).count()
    shiftimings = ShiftTimings.objects.all()
    for member in team:
        teamname = member.Teamname

    if team.exists():
        Context = {'teamcount': teamcount, 'shiftimings': shiftimings, 'username': request.session['username'],
                   'team': team, 'work': work, 'teamname': teamname}
    else:
        Context = {'norecords': 'nodata'}

    return  render(request,'index.html', context=Context)

def report(request):

    Context = {'username': request.session['username']}
    return render(request, 'Report.html', context=Context)

def noticeperiod(request):

    team = Team.objects.filter(Q(Tlmail=request.session['email']))
    teamcount = Team.objects.filter(Tlmail=request.session['email']).count()

    for member in team:
        teamname = member.Teamname
    Context = {'teamcount': teamcount, 'team': team, 'username': request.session['username'], 'teamname': teamname}
    return render(request,'noticeperiod.html',context=Context)

def getreportsatya(request):
    logger.info(str(f'Satya consolidate report generated report on') + str(datetime.now()))
    getreport = GetReport(fromdate=request.POST['reportfromdate'], todate=request.POST['reporttodate'],email="")
    response = getreport.Generate_satya()
    return response

def getreport(request):
    email = request.session['email']
    logger.info(str(f'{email} generated report on') + str(datetime.now()))
    getreport = GetReport(fromdate=request.POST['reportfromdate'], todate=request.POST['reporttodate'],email=request.session['email'])
    response = getreport.Generate()
    return response

def generatereport(request):
    logger.info(str("Satya generated report on ") + str(datetime.now()))
    target_date_str = request.POST["simpleDateInput"]
    generatereport = GenerateReport(target_date_str)
    response = generatereport.Generate()
    Context = {'username': request.session['username']}
    return response

def processteamform(request):

    team = Team.objects.filter(Q(Tlmail=request.session['email']))
    input_date_str = request.POST["simpleDateInput"]
    input_date = datetime.strptime(input_date_str, "%d/%m/%Y").date()
    teamcount = Team.objects.filter(Tlmail=request.session['email']).count()

    work = workon.objects.all()

    team = Team.objects.filter(Q(Tlmail=request.session['email']))
    for member in team:
        teamname = member.Teamname

    try:
        for member in team:

            dailyattendance = DailyAttendance()
            dailyattendance.attendancedate = input_date
            dailyattendance.Employeeno = member.Employeeno
            dailyattendance.Employeename = member.Employeename
            dailyattendance.Teamname = member.Teamname
            dailyattendance.Sapid = member.Sapid
            dailyattendance.Shift = request.POST[str(member.Employeeno)+str("shift")]
            dailyattendance.remarks = request.POST[str(member.Employeeno)+str("remarks")]
            dailyattendance.Tlname = member.Tlname
            dailyattendance.Tlmail = member.Tlmail
            dailyattendance.worktype = request.POST[str(member.Employeeno)]
            dailyattendance.save()
            logger.info(str(dailyattendance.Employeeno) + " savedsuccessfully on " + str(datetime.now()))

            try:
                team_record = Team.objects.get(Employeeno=member.Employeeno)
                new_shift_timing=ShiftTimings.objects.get(shifttiming=request.POST[str(member.Employeeno) + str("shift")])
                team_record.Shift =new_shift_timing
                team_record.save()
                logger.info(str(dailyattendance.Employeeno) + "new timings updated successfully on " + str(datetime.now()))
            except Exception as e:
                logger.info(e)

            team = DailyAttendance.objects.filter(Q(Tlmail=request.session['email'])).order_by('-attendancedate')[:500]
            for member in team:
                input_datetime = datetime.strptime(str(member.attendancedate), "%Y-%m-%d %H:%M:%S%z")
                member.attendancedate = input_datetime.strftime("%d-%m-%Y")

        Context = {'teamcount':teamcount, 'message': 'DataSaved','username': request.session['username'], 'team': team, 'work': work, 'teamname': teamname}

    except Exception as e:

        logger.info("Exception " + str(e) + str(datetime.now()))

        team = DailyAttendance.objects.filter(Q(Tlmail=request.session['email'])).order_by('-attendancedate')[:500]
        for member in team:
            input_datetime = datetime.strptime(str(member.attendancedate), "%Y-%m-%d %H:%M:%S%z")
            member.attendancedate = input_datetime.strftime("%d-%m-%Y")

        e= "User data is already provided for the date"

        Context = {'message': e, 'username': request.session['username'], 'team': team, 'work': work, 'teamname': teamname}

    return render(request, 'read.html', context=Context)





