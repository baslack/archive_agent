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
import openpyxl, os, subprocess, shutil

__author__ = 'Benjamin A. Slack, iam@niamjneb.com'
__version__ = '0.0.0.1'

kBaseJobsPath = '/Volumes/JobsA'
kBaseDisksPath = '/Volumes/SGB-TITAN/_ReadyForBackup'
kWorkingPath = '/Volumes/SGB-TITAN/_ReadyForBackup'
kJobFolderPrefix = '/Jobs'
kTrashFolderPrefix = '/Trash'
kDiscFolderPrefix = '/Disc'
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

    if type(job) != type(''):
        job = str(job)
    folderNumber = job[len(job)-1:len(job)] #get the last number of the job
    return kBaseJobsPath+kJobFolderPrefix+folderNumber+kSep+job



def generate_working_url(job):
    """

    :param job: the job number to generate a working folder for
    :return: the string of the URL of that working folder
    """
    if type(job) != type(''):
        job = str(job)
    return kWorkingPath+kSep+job

def generate_disc_url(disc):
    """

    :param disc: the disc number to generate path to string for
    :return: the path to string
    """
    if type(disc) != type(''):
        disc = str(disc)
    return kBaseDisksPath+kDiscFolderPrefix+disc


def dump_trash(job):
    """

    :param job: The job number to dump the trash of
    :return:
    """
    deleteMe = generate_job_url(job)+kTrashFolderPrefix
    #shutil.rmtree(deleteMe)
    return deleteMe


def copy_job(job):
    """

    :param job: The job number to copy to the work folder
    :return: copy_url
    """
    copyMeFrom = generate_job_url(job)
    copyMeTo = generate_working_url(job)
    command = 'ditto'
    do_this = [command, copyMeFrom, copyMeTo]
    #subprocess.call(do_this)
    print(do_this)
    return copyMeTo


def zip_job(job):
    """

    :param job: The job number to zip using Archive Utility
    :return: zip file path
    """

    command = ['ditto']
    args = ['-c','-k', '--sequesterRsrc', '--keepParent']
    working_path = [generate_working_url(job)]
    zip_path = [kWorkingPath + '/' + str(job) + '.zip']
    do_this = command + args + working_path + zip_path

    #subprocess.call(do_this)
    print(do_this)
    return zip_path[0]


def bucket_job(zip):
    """

    :param zip: zip to put into disc folder
    :return: disc number ZIP was put in

    1. check zip size
    2. check for disc folders
        a. found
            i. for each
                1. check size
                2. check fit
                    a. fits
                        i. move into folder
                    b. doesn't fit
                        i. next folder
                3. no fits
                    a. create folder
                    b. move into folder

        b. not found
            i. create folder
            ii. copy into folder

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
    print(generate_job_url(17545))
    print(generate_working_url(17545))
    print(generate_disc_url(1534))
    print(dump_trash(15687))
    print(copy_job(68975))
    print(zip_job(566843))