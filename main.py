"""
Uses an Excel file to compress and sort jobs into CD folders for burning

Logic Flow:
I. Get a list of jobs from the excel file

II. With that List do the following:

    A. Goto the folder for that job.

    B. Empty the Trash folder

    C. Copy that job folder to a working folder.

    D. ZIP the contents of that copied folder with the OSX Archive Utility

    E. Sort the ZIP into a Disk Folder
        1. Disk folder exists?
            i. Yes.
                a. Disk folder is full? Goto Next Folder
                b. Next folder doesn't exist? Create it and put ZIP in folder. Note folder.
                c. Next folder exists? Is full? Next Folder. Else, place ZIP, note folder.
            ii. No.
                a. Create Folder.
                b. Place ZIP.
                c. Note Folder.

    F. Delete the Working Copy

    G. Delete the Contents of the Job Folder

    H. Add a folder, noting the Disk Folder where the ZIP was placed

    I. Append a Disk number to that job's entry in the Excel document


"""
import openpyxl, os

__author__ = 'Benjamin A. Slack, iam@niamjneb.com'
__version__ = '0.0.0.1'

kBaseJobsPath = '/Volumes/JobsA'
kBaseDisksPath = '/SGB-TITAN/_ReadyForBackup'
kWorkingPath = '/SGB-TITAN/_ReadyForBackup'
kJobFolderPrefix = '/Jobs'
kFullSize = '4294967296'  # 4GB
kFileName = 'test.xlsx'
kPath = os.path.expanduser('~')
kSep = '/'
kURL = kPath + kSep + kFileName


def get_list(url):
    """

    :param url: The location of the excel file we're using
    :return: the sequence of job numbers retrieved from it
    """

    myList = list()
    wb = openpyxl.load_workbook(url)
    ws = wb.active
    r,c = 1,1
    while ws.cell(row=r, column=c).value:
        myList.append(ws.cell(row=r, column=c).value)
        r = r + 1
    return myList

def generate_job_url(job):
    """

    :param job: the
     job number to generate the URL for
    :return: the string of the URL to the job on the server
    """


def generate_working_url(job):
    """

    :param job: the job number to generate a working folder for
    :return: the string of the URL of that working folder
    """

def generate_disc_url(disc):
    """

    :param disc: the disc number to generate path to string for
    :return: the path to string
    """


def dump_trash(job):
    """

    :param job: The job number to dump the trash of
    :return:
    """


def copy_job(job):
    """

    :param job: The job number to copy to the work folder
    :return: copy_url
    """


def zip_job(job):
    """

    :param job: The job number to zip using Archive Utility
    :return: zip filename
    """


def bucket_job(zip):
    """

    :param zip: zip to put into disc folder
    :return: disc number ZIP was put in
    """

    def check_disc(disc):
        """
        :param disc: the disc number of the folder to check for size
        :return: full or not
        """

    def create_disc(last_disc):
        """
        :param last_disc: the disc number of the last full disc
        :return the number of the new disc
        """


def dump_copy(copy_url):
    """

    :param copy_url: the url of the temp folder to delete
    :return:
    """


def dump_job(job):
    """

    :param job: the job number to dump the contents on
    :return:
    """


def tag_job(job, disc):
    """

    :param job: the job number to add the disc tag folder to
    :param disc: the disc number for the tag
    :return:
    """


def tag_excel(url, job, disc):
    """

    :param url: location of the excel file
    :param job: job number to lookup in the excel file
    :param disc: disc number to enter in the excel file
    :return:
    """


if __name__ == "__main__":
    print(get_list(kURL))