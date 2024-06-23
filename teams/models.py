from django.db import models

class ShiftTimings(models.Model):
    shifttiming=models.CharField(max_length=50, primary_key=True)

    class Meta:
        verbose_name_plural = "Shifttimings"

    def __str__(self):
        return self.shifttiming

class DailyAttendance(models.Model):
    attendancedate = models.DateTimeField()
    Employeeno = models.IntegerField()
    Employeename = models.CharField(max_length=100, null=True, blank=True)
    Teamname = models.CharField(max_length=100, null=True, blank=True)
    Sapid = models.CharField(max_length=50, null=True, blank=True)
    Shift = models.CharField(max_length=100, null=True, blank=True)
    Tlname = models.CharField(max_length=100, null=True, blank=True)
    Tlmail = models.CharField(max_length=100, null=True, blank=True)
    worktype = models.CharField(max_length=100, null=True, blank=True)
    remarks = models.CharField(max_length=100, null=True, blank=True)

    class Meta:
        verbose_name_plural = "AttendanceTracker"
        # Define a UniqueConstraint for the composite primary key
        constraints = [
            models.UniqueConstraint(fields=['attendancedate', 'Employeeno'], name='unique_attendance')
        ]


class teamnamedetail(models.Model):
    teamdetail = models.CharField(max_length=100, primary_key=True)

    def __str__(self):
        return self.teamdetail

    class Meta:
        verbose_name_plural = "teamname"


class workon(models.Model):
    worktype = models.CharField(max_length=100, null=True, blank=True)

    class Meta:
        verbose_name_plural = "WorkType"

class Login(models.Model):
    username = models.CharField(max_length=100, primary_key=True)
    password = models.CharField(max_length=100, null=True, blank=True)

    def __str__(self):
        return self.username

    class Meta:
        verbose_name_plural = "Login"


# Create your models here.
class Team(models.Model):
    Employeeno = models.IntegerField(primary_key=True)
    Employeename = models.CharField(max_length=100, null=True, blank=True)
    #Teamname = models.CharField(max_length=100, null=True, blank=True)
    Teamname = models.ForeignKey(teamnamedetail, on_delete=models.CASCADE, to_field='teamdetail', null=True, blank=True)
    Sapid = models.CharField(max_length=50, null=True, blank=True)
    Tlname = models.CharField(max_length=100, null=True, blank=True)
    Shift = models.ForeignKey(ShiftTimings, on_delete=models.CASCADE, to_field='shifttiming', null=True, blank=True)
    Tlmail = models.ForeignKey(Login, on_delete=models.CASCADE, to_field='username', null=True, blank=True)

    class Meta:
        verbose_name_plural = "TeamDetails"
