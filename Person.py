class Person():
    def __init__(self,name, gender, phone_number, ssn, start_date):
        self.name = name
        self.gender = gender
        self.phone_number = phone_number
        self.ssn = ssn
        self.start_date = start_date
    def leave(self, leave_date):
        self.leave_date = leave_date

    def isLeft(self):
        if self.leave_date != None:
            return True
        else:
            return False
