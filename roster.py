import pandas, openpyxl
from openpyxl import load_workbook


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
        self.workbook = load_workbook(filename)
        return

    def get_student_names(self):
        """
        Get student names.

        Returns:
            List of student names

        """
        return self.workbook.get_student_names().list()

    def get_student(self, name):
        """
        Get student.

        Args:
            name: Name of the student.

        Returns:
            This is a description of what is returned.

        """
        return self.workbook.get_student(name).str()

    def class_average(self):
        """
        Get class average.

        Returns:
            int class average.

        """
        return self.workbook.get_class_average().int()

    def delete_student(self, name):
        """
        Remove a student from the workbook.

        Args:
            filename: Name of the student.

        """
        self.workbook.get_student_names(name)
        return

    def save(self, filename):
        """
        Write workbook to file.

        Args:
            filename: Name of the workbook file.

        Returns:
            This is a description of what is returned.

        """
        self.workbook.save_workbook()
        return
