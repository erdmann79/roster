import pandas, openpyxl
from openpyxl import load_workbook
from pprint import pprint
from itertools import islice


class Student:

    def __init__(self, my_dataframe):
        self.my_dataframe = my_dataframe

class Roster(object):
    """
    Summery of Roster Class.

    Mrs. Jones, a teacher at Billings Elementary School, has been asked by the administration to
    track her students' grades in a roster stored in an Excel file. Unfortunatly, Mrs. Jones had a
    bad experience with Excel as a child and as solemnly sworn never to use Excel again. Your task
    is to create a Python class for reading and manipulating a provided Excel file so Mrs. Jones
    doesn't haven't to use Excel to successfully record her students' grades.

    """

    def __init__(self, filename):
        """
        Init SampleClass with blah.

        Args:
            filename: Name of the workbook file.
        """
        self.filename = filename
        self.workbook = None
        self.roster_sheet = None
        self.first_name_column = None
        self.last_name_column = None

    def __enter__(self):
        """Open the workbook."""
        self.workbook = load_workbook(self.filename)
        self.roster_sheet = self.workbook["Roster"]
        return self

    def __exit__(self, exc_type, exc_value, traceback):
        """Close the workbook."""
        self.workbook.close()

    def get_student_names(self):
        """
        Get student names.

        Returns:
            List of student names

        """
        roster_sheet = self.workbook["Roster"]
        student_name_list = []
        for (first, last) in roster_sheet.iter_rows(
            min_row=2, min_col=2, max_col=3, max_row=1000, values_only=True
        ):
            if not first and not last:
                break
            student_name_list.append(first + " " + last)

        return student_name_list

    def get_student(self, name):
        """
        Get student.

        Args:
            name: Name of the student.

        Returns:
            This is a description of what is returned.

        """

        

        

        if name in self.get_student_names():
            student = {}
            assingment_results = []
        for (id_number, first, last) in self.roster_sheet.iter_rows(
            min_row=2, min_col=1, max_col=3, values_only=True
        ):
            if not first and not last:
                break
            if name == first + " " + last:
                student["id"] = id_number
                break




        for i, student_sheet in enumerate(self.workbook):
            if i == 0:
                continue
            if name == student_sheet["B2"].internal_value:
                #student["student_sheet"] = student_sheet.title

                print('*************0*')

                data = student_sheet.values
                # Get the first line in file as a header line
                
                next(data)
                next(data)
                next(data)
                next(data)
                columns = next(data)[0:]

                print(columns)
                df = pandas.DataFrame(data, columns=columns)
                df = df.rename(columns={"Grade": "grades"})
                #student['id'] = id_number
                student['grades'] = Student(df['grades'])

                #
                #df["student_sheet"] = student_sheet.title
                #student["student_sheet_dataframe"] = df
                student = df
                #student['grades'] = df['grades']
                #print(df['Grade'])
                #print(student["student_sheet_dataframe"])

                #print(student["student_sheet_dataframe"]["grades"])

                #print('*************1*')
                #student["student_sheet_dataframe"] = pandas.DataFrame(student_sheet.values)
                #print(student["student_sheet_dataframe"])

                

                # df = pandas.read_excel('29.xlsx', engine='openpyxl',index_col=None)

                # data = student_sheet.values
                # cols = next(data)[1:]
                # data = list(data)
                # idx = [r[0] for r in data]
                # data = (islice(r, 1, None) for r in data)

                # df = pandas.DataFrame(data, index=idx, columns=cols)
                # df = pandas.DataFrame(student_sheet)

                # my_series = df['Price'].squeeze()
                # print(my_series)

                # print('**************')

                #for (assingment_result,) in student_sheet.iter_rows(min_row=6, min_col=2, max_col=2, values_only=True):
                #    assingment_results.append(assingment_result)
                #student["grades"] = pandas.Series(assingment_results)




                #print(dir(student_sheet.iter_rows))

                #
                #student["grades"] = pandas.Series(
                #    student_sheet.iter_rows(
                #        min_row=6, min_col=1, max_col=2, values_only=True
                #    )[0]
                #)
                print(student)
                print('*1**************')

                return student

    def class_average(self):
        """
        Get class average.

        Returns:
            int class average.

        #    self.assertTrue(roster_obj.class_average() == 614.1 / 7)

        """
        student_avg_list = []
        for i, student_sheet in enumerate(self.workbook):
            if i == 0:
                continue
            assingment_results = []
            for (assingment_result,) in student_sheet.iter_rows(
                min_row=6, min_col=2, max_col=2, values_only=True
            ):
                assingment_results.append(assingment_result)

            student_avg_list.append(pandas.Series(assingment_results).mean())

        class_average = sum(student_avg_list) / len(student_avg_list)
        return class_average

    def delete_student(self, name):
        """
        Remove a student from the workbook.

        Args:
            filename: Name of the student.
        """
        print(len(self.get_student_names()))
        student_sheet = self.get_student(name)["student_sheet"]
        self.workbook.remove_sheet(student_sheet)
        student_list = self.get_student_names()
        shift_register = 2
        for i, student_name in enumerate(student_list):
            row = i + shift_register
            id_number = i + shift_register - 1
            if student_name == name:
                print("remove", name, "row", row, "id", id_number)
                self.roster_sheet.delete_rows(row)
                shift_register = 1
                continue
            print(row, id_number, student_name)
            if shift_register == 1:
                self.roster_sheet[f"A{row}"] = id_number
                student_sheet = self.get_student(student_name)["student_sheet"]
                student_sheet["B1"] = id_number
                student_sheet.title = f"Student_{id_number}"
                print(student_sheet.title)
        print(len(self.get_student_names()))

    def save(self, filename=None):
        """
        Write workbook to file.

        Args:
            filename: Name of the workbook file.

        """
        if not filename:
            filename = self.filename
        print(filename)

        self.workbook.save(filename)


if __name__ == "__main__":
    with Roster("Jones_2019.xlsx") as roster_obj:
        
        #student_names = roster_obj.get_student_names()
        #print(student_names)


        john = roster_obj.get_student("Johnny Carson")

        print("***** Main ************************")
        print()
        print('*********before***************')
        print(john)
        print(john["grades"][6],85)
        print(roster_obj.class_average(),614.1 / 7)        
        for assignment, grade in [(3, 90), (6, 94), (9, 92)]:
            john["grades"][assignment] = grade
        print()
        print('*********after*************')
        print(john["grades"][6],94)
        print(roster_obj.class_average(),614.1 / 7)

        john = roster_obj.get_student("Johnny Carson")
        print(john)
        print(john["grades"][6],94)
        print(roster_obj.class_average(),614.1 / 7)
        print('**********************')




    #    pprint(student_record)
    #    class_avg = roster_obj.class_average()
    #    print(class_avg)
    #    print(614.1 / 7)
#
#    roster_obj.delete_student("Allen Dalton")
#    roster_obj.save("Jones_2019_Reduced.xlsx")
#    workbook_reduced = load_workbook("Jones_2019_Reduced.xlsx")
#
# sheet_names = workbook_reduced.get_sheet_names()
# print(len(sheet_names))
