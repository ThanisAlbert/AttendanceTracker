from django.urls import path
from . import  views

app_name="teams"

urlpatterns = [
    path('', views.login,name='login'),
    path('logout.html',views.logout,name='logout'),
    path('index.html',views.processlogin, name='processlogin'),
    path('home.html',views.home,name='home'),
    path('noticeperiod.html',views.noticeperiod,name='noticeperiod'),
    path('processteamform',views.processteamform,name='teamform'),
    path('read.html',views.readdata,name='read'),
    path('edit.html',views.editattendance,name='editattendance'),
    path('team.html',views.team,name='team'),
    path('updateattendance.html',views.updateattendance,name='updateattendance'),
    path('report.html',views.report,name='report'),
    path('getreport_satya.html',views.getreportsatya,name='getreport_satya'),
    path('getreport.html',views.getreport,name='getreport'),
    path('generatereport.html',views.generatereport,name='generatereport'),
]
