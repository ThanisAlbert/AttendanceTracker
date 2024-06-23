from datetime import datetime

from teams.models import DailyAttendance

class Shift:

    def __init__(self,attendancedate):
        self.attendancedate = datetime.strptime(attendancedate, "%d/%m/%Y")

    def GetShiftTimings(self):

        try:
            shifttimings = set()

            attendance_records = DailyAttendance.objects.filter(attendancedate=self.attendancedate)

            for record in attendance_records:
                if record.Shift !="":
                    shifttimings.add(str(record.Shift))
                else:
                    print("testing")

            print(shifttimings)

            sorted_times = sorted(shifttimings, key=self.sort_key)

            am_shift=[]
            pm_shift =[]
            pm12_shift=[]

            consold_shift=[]

            for time in sorted_times:
                if "AM" in str(time).split("to")[0].upper():
                    am_shift.append(time)

            for time in sorted_times:
                if "PM" in str(time).split("to")[0].upper() and "12" not in str(time).split("to")[0].upper():
                    pm_shift.append(time)

                if "PM" in str(time).split("to")[0].upper() and "12" in str(time).split("to")[0].upper():
                    pm12_shift.append(time)

            consold_shift = am_shift+pm12_shift+pm_shift

            return consold_shift

        except Exception as e:
            print("Exception occured")
            print(e)

        return shifttimings

    def sort_key(self,time_str):
        if "." in time_str.split("to")[0]:
            return int(time_str.split('.')[0])
        else:
            return int(time_str.split(':')[0])















