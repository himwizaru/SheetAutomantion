"""Script to Automate Whatsapp Messages from Google Sheets"""
import datetime, json, requests
from googleapiclient.discovery import build
from google.oauth2 import service_account
from decouple import config

class AutomateSheets:
    """Class to Automate the Google Sheets Messages"""
    def __init__(self):
        self.delay = int(config('DELAY'))
        self.retry_threshold = int(config('RETRYTHRESHOLD'))
        self.endpoint = "https://api.gupshup.io/sm/api/v1/msg"
        self.spreadsheet_id = config('SPREADSHEET_ID')
        self.spreadsheet_name = config('SPREADSHEETNAME')
        self.scopes = ['https://www.googleapis.com/auth/spreadsheets']
        self.keypath = './' + config('APIFILE')
        self.creds = service_account.Credentials.from_service_account_file(self.keypath, \
        scopes=self.scopes)
        self.service = build('sheets', 'v4', credentials=self.creds)
        self.sheet = self.service.spreadsheets()
        self.mapping = {
            "parent_first_name" : 1, "parent_last_name" : 2, "child_first_name" : 3,
            "child_last_name" : 4, "phone_number" : 5, "course_name" : 6,
            "class_date" : 7, "class_time" : 8, "reminder_date" : 9, "reminder_time": 10,
            "message_date" : 11, "message_time" : 12, "class_link" : 13, "time_zone" : 14,
            "message_sent" : 15, "message_sent_time" : 16, "message_retry" : 17,
            "reminder_sent" : 18, "reminder_sent_time" : 19, "reminder_retry" : 20
        }
        self.day_mapping = ['Monday', 'Tuesday', 'Wednesday', 'Thursday',
                            'Friday', 'Saturday', 'Sunday'] 
        self.month_mapping = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug",
                              "Sep", "Oct", "Nov", "Dec"]
        self.time_format = config("TIMEFORMAT")

    # pylint: disable=R0201
    def format_time(self, time_string, data_date):
        """Use for time formatting"""
        time_array = time_string.split(" ")
        hours, mins = time_array[0].split(":")
        hours = int(hours)
        if time_array[-1] == self.time_format:
            if hours != 12:
                hours += 12
        else:
            if hours == 12:
                hours -= 12
        data_date_array = [int(x) for x in data_date.split("/")]
        d = self.get_custom_datetime_object(time_string, data_date_array[0], \
            data_date_array[1], data_date_array[2])
        modified_time = str(d.year)
        if len(str(d.month)) == 1:
            modified_time += ("0" + str(d.month))
        else:
            modified_time += str(d.month)
        if len(str(d.day)) == 1:
            modified_time += ("0" + str(d.day))
        else:
            modified_time += str(d.day)
        if len(str(d.hour)) == 1:
            modified_time += ("0" + str(d.hour))
        else:
            modified_time += str(d.hour)
        if len(str(d.minute)) == 1:
            modified_time += ("0" + str(d.minute))
        else:
            modified_time += str(d.minute)
        return int(modified_time)

    # pylint: disable=R0201
    def reverse_time_formatting(self, hrs, mins, time_zone):
        """Reverse Fromat the Time String"""
        return_string = ""
        if hrs >= 12:
            if hrs != 12:
                hrs -= 12
            return_string += (" PM " + time_zone)
        else:
            if hrs == 0:
                hrs += 12
            return_string += (" AM " + time_zone)
        hrs = str(hrs)
        mins = str(mins)
        if len(hrs) == 1:
            hrs = "0" + hrs
        if len(mins) == 1:
            mins = "0" + mins
        return hrs + ":" + mins + return_string

    # pylint: disable=R0201
    def get_custom_datetime_object(self, time_string, date, month, year):
        """Returns Cur=stom Date Time Object"""
        time_string, time_stamp = time_string.split(" ")
        hrs, mins = [int(x) for x in time_string.split(":")]
        if time_stamp == self.time_format:
            if hrs != 12:
                hrs += 12
        else:
            if hrs == 12:
                hrs -= 12
        d = datetime.datetime(year, month, date, hrs, mins)
        return d

    # pylint: disable=R0201
    def get_time(self, data):
        """Returns the Current time"""
        curr_time = datetime.datetime.now()
        data_date = data[self.mapping["class_date"]]
        data_date_array = data_date.split("/")
        data_date = int(data_date_array[0])

        ahead_time = curr_time + datetime.timedelta(minutes=self.delay)
        curr_time_str = str(curr_time.year)
        if len(str(curr_time.month)) == 1:
            curr_time_str += ("0" + str(curr_time.month))
        else:
            curr_time_str += str(curr_time.month)
        if len(str(curr_time.day)) == 1:
            curr_time_str += ("0" + str(curr_time.day))
        else:
            curr_time_str += str(curr_time.day)
        if len(str(curr_time.hour)) == 1:
            curr_time_str += ("0" + str(curr_time.hour))
        else:
            curr_time_str += str(curr_time.hour)
        if len(str(curr_time.minute)) == 1:
            curr_time_str += ("0" + str(curr_time.minute))
        else:
            curr_time_str += str(curr_time.minute)

        ahead_time_str = str(ahead_time.year)
        if len(str(ahead_time.month)) == 1:
            ahead_time_str += ("0" + str(ahead_time.month))
        else:
            ahead_time_str += str(ahead_time.month)
        if len(str(ahead_time.day)) == 1:
            ahead_time_str += ("0" + str(ahead_time.day))
        else:
            ahead_time_str += str(ahead_time.day)
        if len(str(ahead_time.hour)) == 1:
            ahead_time_str += ("0" + str(ahead_time.hour))
        else:
            ahead_time_str += str(ahead_time.hour)
        if len(str(ahead_time.minute)) == 1:
            ahead_time_str += ("0" + str(ahead_time.minute))
        else:
            ahead_time_str += str(ahead_time.minute)

        d = self.get_custom_datetime_object(data[self.mapping["class_time"]], \
            int(data_date_array[0]), int(data_date_array[1]), int(data_date_array[2]))
        if data[self.mapping["time_zone"]] == "EST":
            d = d - datetime.timedelta(hours=9, minutes=30)
        time_zone_time = self.reverse_time_formatting(d.hour, d.minute, \
            data[self.mapping["time_zone"]])
        get_date_string = str(d.day) + " " + self.month_mapping[d.month - 1] + " " + \
            str(d.year)   
        curr_day = str(d.day)
        if len(curr_day) == 1:
            curr_day = "0" + curr_day
        curr_month = str(d.month)
        if len(curr_month) == 1:
            curr_month = "0" + curr_month
        curr_year = str(d.year)
        get_day_format = curr_day + " " + curr_month + " " + curr_year
        get_day = datetime.datetime.strptime(get_day_format, '%d %m %Y').weekday()
        get_day = self.day_mapping[get_day]
        return [int(curr_time_str), int(ahead_time_str), \
            get_day, get_date_string, time_zone_time]

    # pylint: disable=R0201
    def get_message(self, child_name, course, class_date, \
        class_day, class_time, meet_link, message_type):
        """Returns The Modified Message"""
        message = ""
        if message_type == 1:
            message = f"*Gentle Reminder, class starts in 30 mins* \n\nHey Parent,\n\n*{child_name}'s* class is booked:\n📝: *{course}* Classes for Kids\n📅: *{class_date}* ({class_day}), *{class_time}*\n\nInstructions to join:\n- Join using Mobile: Download Zoom\n- Join using Computer/ desktop:\n{meet_link}\n\nHappy Learning!\n\nTeam Wizaru\nwww.wizaru.com"
        else:
            message = f"Hey Parent,\n\n*{child_name}'s* class is booked:\n📝: *{course}* Classes for Kids\n📅: *{class_date}* ({class_day}), *{class_time}*\n\nInstructions to join:\n- Join using Mobile: Download Zoom\n- Join using Computer/ desktop:\n{meet_link}\n\nHappy Learning!\n\nTeam Wizaru\nwww.wizaru.com"

        return message

    def notify(self, data, row, message_type, curr_date, curr_day, time_zone_time):
        """ Send the POST request to the GupShup API and update the fields in the Google Sheets"""
        headers = {
            'Cache-Control' : 'no-cache',
            'Content-Type' : 'application/x-www-form-urlencoded',
            'apikey' : config('APIKEY')
        }
        msg = ""
        if message_type == "reminder":
            msg = self.get_message(data[self.mapping["child_first_name"]], \
                data[self.mapping["course_name"]], curr_date, curr_day, \
                    time_zone_time, data[self.mapping["class_link"]], 1)
        else:
            msg = self.get_message(data[self.mapping["child_first_name"]], \
                data[self.mapping["course_name"]], curr_date, curr_day, \
                    time_zone_time, data[self.mapping["class_link"]], 0)
        
        msg_string = {'type':'text', 'text': msg}
        destination_phone = data[self.mapping["phone_number"]].split("-")
        send_data = {
            'channel' : 'whatsapp',
            'source' : config('SRCPH'),
            'destination' : destination_phone[0] + destination_phone[-1],
            'message' : json.dumps(msg_string),
            'src.name' : config('BOTNAME')
        }
        req = requests.post(self.endpoint, data=send_data, headers=headers)
        if req.status_code == 200:
            rng_sent = ""
            time = datetime.datetime.now()
            time = str(time.hour) + ":" + str(time.minute)
            if message_type == "reminder":
                rng_sent = "Class Schedule" + f"!S{row+1}"
            else:
                rng_sent = "Class Schedule" + f"!P{row+1}"
            self.sheet.values().update(spreadsheetId=self.spreadsheet_id, range=rng_sent, \
            valueInputOption="USER_ENTERED", body={"values":[[1, time]]}).execute()
        else:
            val = ""
            rng_sent = ""
            if message_type == "reminder":
                val = int(data[self.mapping["reminder_retry"]]) + 1
                rng_sent = "Class Schedule" + f"!S{row+1}"
            else:
                val = int(data[self.mapping["message_retry"]]) + 1
                rng_sent = "Class Schedule" + f"!P{row+1}"
            self.sheet.values().update(spreadsheetId=self.spreadsheet_id, range=rng_sent, \
            valueInputOption="USER_ENTERED", body={"values":[[str(val)]]}).execute()

    def worker(self):
        """ Main Function to be called after the fixed delay"""
        result = self.sheet.values().get(spreadsheetId=self.spreadsheet_id, \
        range=self.spreadsheet_name).execute()
        data = result.get('values', [])

        for i in range(1, len(data)):
            curr_time = self.get_time(data[i])
            curr_time_int = curr_time[0]
            curr_time_ahead = curr_time[1]

            # For Sending the Reminder
            reminder_time_int = self.format_time(data[i][self.mapping["reminder_time"]], \
                data[i][self.mapping["reminder_date"]])
            if int(data[i][self.mapping["reminder_retry"]]) >= self.retry_threshold:
                continue

            if int(data[i][self.mapping["reminder_sent"]]) == 0 and \
                ((curr_time_ahead >= reminder_time_int and \
                reminder_time_int >= curr_time_int) or \
                (curr_time_int > reminder_time_int and \
                    int(data[i][self.mapping["reminder_retry"]]) > 0)):
                self.notify(data[i], i, "reminder", curr_time[3], curr_time[2], \
                    curr_time[4])

            # For Sending the Message
            message_time_int = self.format_time(data[i][self.mapping["message_time"]], \
                data[i][self.mapping["message_date"]])
            if int(data[i][self.mapping["message_retry"]]) >= self.retry_threshold:
                continue

            if int(data[i][self.mapping["message_sent"]]) == 0 and \
                ((curr_time_ahead >= message_time_int and \
                message_time_int >= curr_time_int) or \
                (curr_time_int > message_time_int and \
                int(data[i][self.mapping["message_retry"]]) > 0)):
                self.notify(data[i], i, "message", curr_time[3], curr_time[2], \
                    curr_time[4])

OBJ = AutomateSheets()
OBJ.worker()