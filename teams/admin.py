from django.contrib import admin
from django.contrib.admin.models import LogEntry
from import_export.admin import ImportExportModelAdmin
from import_export import resources
from teams.models import Team, Login, workon, DailyAttendance, ShiftTimings, teamnamedetail

class teamresource(resources.ModelResource):
    class Meta:
        model=Team
        exclude = ('id',)
        import_id_fields = ('Employeeno',)

class teamadmin(ImportExportModelAdmin):
    resource_class = teamresource
    search_fields = ['Employeeno']
    list_display = ['Employeeno', 'Employeename', 'Teamname', 'Sapid', 'Shift', 'Tlname','Tlmail']


class loginresource(resources.ModelResource):
    class Meta:
        model=Login
        exclude = ('id',)
        import_id_fields = ('username',)

class loginadmin(ImportExportModelAdmin):
    resource_class = loginresource
    search_fields = ['username']
    list_display = ['username', 'password']


class workonresource(resources.ModelResource):
    class Meta:
        model=workon
        exclude = ('id',)
        import_id_fields = ('worktype',)

class workadmin(ImportExportModelAdmin):
    resource_class = workonresource
    search_fields = ['worktype']
    list_display = ['worktype']


class attendanceresource(resources.ModelResource):
    class Meta:
        model=DailyAttendance


class attendanceadmin(ImportExportModelAdmin):
    resource_class = attendanceresource
    search_fields = ['Employeeno']
    list_display = ['attendancedate','Employeeno','Employeename','Teamname','Sapid','Shift','Tlname','Tlmail','worktype','remarks']


admin.site.site_header = 'Redserv'
admin.site.register(teamnamedetail)
admin.site.register(ShiftTimings)
admin.site.register(Team,teamadmin)
admin.site.register(Login,loginadmin)
admin.site.register(workon,workadmin)
admin.site.register(DailyAttendance,attendanceadmin)

