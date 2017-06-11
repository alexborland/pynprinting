"""Module for working with NPrinting NSQ files"""
import xml.etree.ElementTree as ET
import csv
import os

source_directory = "X:/NPrinting/NSQFiles"
destination_directory = "output"

def prop_to_dict(property_list):
    """Returns a dict {propertyname : property object}"""
    output = {}
    for p in property_list:
        output[p.attrib["name"]] = p
    return output

def dump_recips_from_directory(source_dir, dest_dir):
    for filename in os.listdir(source_dir):
        if filename.endswith(".nsq"):
            print(f"Pulling recipients from {filename}...")
            out_file_name = f"{dest_dir}/{filename}.csv".replace(".nsq","")
            NSQ(f"{source_dir}/{filename}").dump_user_import_file(out_file_name)

class NSQ(object):
    def __init__(self, filename):
        raw_text = "<?xml version=\"1.0\" encoding=\"UTF-8\"?>"
        with open(filename) as f:
            f.readline()
            raw_text += f.read()
        xml = ET.XML(raw_text)
        properties = prop_to_dict(xml[0][0])

        self.ID = properties["ID"].text
        self.name = properties["Name"]
        self.description = properties["Description"]
        self.label = properties["Label_"]
        self.db_connections = [c for c in properties["DatabaseConnections"]]
        self.destinations = [d for d in properties["Destinations"]]
        self.linked_fields = [f for f in properties["LinkedFields"]]
        self.conditions = [cond for cond in properties["Conditions"]]
        self.filters = [Filter(f) for f in properties["Filters"]]
        self.roles = [role for role in properties["Roles"]]
        self.users = [User(u) for u in properties["Users"]]
        self.user_imports = [i for i in properties["UserImports"]]
        self.groups = [g for g in properties["Groups"]]
        self.qlik_reports = [r for r in properties["QlikReports"]]
        self.reports = [r for r in properties["Reports"]]
        self.Office_Reports = [Office_Report(r) for r in
                               properties["OfficeReports"]]
        self.tasks = [Task(t) for t in properties["Tasks"]]
        self.jobs = [j for j in properties["Jobs"]]
        self.schedules = [s for s in properties["Schedules"]]

    def get_filter(self, filter_id):
        for i in self.filters:
            if i.ID == filter_id:
                return i
        return 0

    def get_user(self, user_id):
        for u in self.users:
            if u.ID == user_id:
                return u
        return 0

    def task_summary(self):
        result = ""
        for task in self.tasks:
            result += f"Task: {task.label}\n"
            for filtID in task.filters:
                result += f"\tTask filter: {self.get_filter(filtID).fields}\n"
            for userID in task.recipients:
                result += f"\tRecipient: {self.get_user(userID).name} "
                result += f"({self.get_user(userID).email})\n"
                for filtID in self.get_user(userID).filters:
                    for field, value in self.get_filter(filtID).fields.items():
                        result += "\t\t"
                        result += f"{field}: {value}"
                        result += "\n"
        return result

    def user_filter_format(self, user):
        result = ""
        if len(user.filters) == 0:
            return ""
        for f_id in user.filters:
            filter = self.get_filter(f_id)
            for field in filter.fields:
                result += (f"[{field.name}] = {{")
                for value in field.values:
                    result += value.value
                    if len(value.tags()) > 0:
                        result += "<"
                        result += ",".join(value.tags())
                        result += ">"
                    result += ","
                result = result[:len(result) - 1]
                result += "}"
                if len(field.tags()) > 0:
                    result += "<"
                    result += ",".join(field.tags())
                    result += ">"
                result += ", "
        result = result[:len(result) - 2]
        return result
    
    def tasks_containing_user(self, user_id):
        results = []
        for task in self.tasks:
            if user_id in task.recipients:
                results.append(task.label)
        return ', '.join(results)

    def dump_user_import_file(self, output_file):
        with open(output_file, "w", newline='') as f:
            csvwriter = csv.writer(f)
            csvwriter.writerow(["Name"]+["Email"]+["Desc (Tasks)"]+["Filters"])
            for user in self.users:
                csvwriter.writerow([user.name]
                                 + [user.email]
                                 + [self.tasks_containing_user(user.ID)]
                                 + [self.user_filter_format(user)])

class Task(object):
    def __init__(self, element):
        properties = prop_to_dict(element[0])
        self.ID = properties["ID"].text
        self.name = properties["Name"].text
        self.description = properties["Description"].text
        self.label = properties["Label_"].text
        self.db_connection = properties["DatabaseConnectionID"].text
        self.recipients = []
        for i in properties["Recipients"][0][0]:
            if i.attrib["name"] == "Recipients":
                for j in i:
                    for k in j[0]:
                        if k.attrib["name"] == "ReferenceID":
                            self.recipients.append(k.text)
        self.filters = []
        for i in properties["Filters"]:
            self.filters.append(i.text)

class User(object):
    def __init__(self, element):
        properties = prop_to_dict(element[0])
        self.ID = properties["ID"].text
        self.name = properties["Name"].text
        self.description = properties["Description"].text
        self.label = properties["Label_"].text
        self.email = properties["Email"].text
        self.filters = []
        for i in properties["Filters"]:
            self.filters.append(i.text)

class Filter(object):
    def __init__(self, element):
        properties = prop_to_dict(element[0])
        self.ID = properties["ID"].text
        self.name = properties["Name"].text
        self.description = properties["Description"].text
        self.label = properties["Label_"].text
        self.fields = [Field(f) for f in properties["Fields"]]

class Field(object):
    def __init__(self, element):
        properties = prop_to_dict(element[0])
        self.ID = properties["ID"].text
        self.name = properties["Name"].text
        self.description = properties["Description"].text
        self.label = properties["Label_"].text
        source_field_elem = properties["SourceField"][0][0]
        self.source_field = prop_to_dict(source_field_elem)["Name"].text
        if "CheckPossible" in properties:
            if properties["CheckPossible"].text == "True":
                self.verify = True
            else:
                self.verify = False
        else:
            self.verify = False
        if "UserCanUnlock" in properties:
            if properties["UserCanUnlock"].text == "True":
                self.unlock = True
            else:
                self.unlock = False
        else:
            self.unlock = False
        if "Excluded" in properties:
            if properties["Excluded"].text == "True":
                self.excluded = True
            else:
                self.excluded = False
        else:
            self.excluded = False
        if "Lock" in properties:
            if properties["Lock"].text == "True":
                self.lock = True
            else:
                self.lock = False
        else:
            self.lock = False
        # TO DO: add the remaining field-level filter options
        #       - drop
        self.values = [Field_Value(v) for v in properties["Values"]]

    def tags(self):
        result = []
        if self.unlock:
            result.append("unlock")
        if self.verify:
            result.append("verify")
        if self.excluded:
            result.append("excluded")
        if self.lock:
            result.append("lock")
        return result

class Field_Value(object):
    def __init__(self, element):
        properties = prop_to_dict(element[0])
        self.ID = properties["ID"].text
        self.name = properties["Name"].text
        self.description = properties["Description"].text
        self.label = properties["Label_"].text
        self.value = properties["Value"].text
        self.number = properties["Number"].text
        if "IsNumeric" in properties:
            if properties["IsNumeric"].text == "True":
                self.is_numeric = True
            else:
                self.is_numeric = False
        else:
            self.is_numeric = False
        if "Evaluate" in properties:
            if properties["Evaluate"].text == "True":
                self.evaluate = True
            else:
                self.evaluate = False
        else:
            self.evaluate = False
    
    def tags(self):
        result = []
        if self.evaluate:
            result.append("evaluate")
        # if self.is_numeric:
        #     result.append("numeric")
        return result
        
    def get_value(self):
        if self.is_numeric:
            return self.number
        else:
            return self.value

class Office_Report(object):
    def __init__(self, element):
        properties = prop_to_dict(element[0])
        self.ID = properties["ID"].text
        self.name = properties["Name"].text
        self.description = properties["Description"].text
        self.label = properties["Label_"].text
        self.db_connection = properties["DatabaseConnectionID"].text
        self.report_type = properties["ReportType"].text
        self.template = properties["Template"].text


dump_recips_from_directory(source_directory, destination_directory)

