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
import openpyxl, os, time, subprocess, shutil

__author__ = 'Benjamin A. Slack, iam@niamjneb.com'
__version__ = '0.0.0.1'

#kBaseJobsPath = '/Volumes/JobsA'
kBaseJobsPath = os.path.expanduser('~/tmp/JobsA')
#kBaseDisksPath = '/Volumes/SGB-TITAN/_ReadyForBackup'
kBaseDisksPath = os.path.expanduser('~/tmp/Staging')
#kWorkingPath = '/Volumes/SGB-TITAN/_ReadyForBackup'
kWorkingPath = os.path.expanduser('~/tmp/Staging')
kJobFolderPrefix = '/Jobs'
kTrashFolderPrefix = '/Trash'
kDiscFolderPrefix = 'Disc'
kFullSize = 4294967296  # 4GB
kFileName = 'test.xlsx'
kPath = os.path.expanduser('~')
kSep = '/'
#kURL = kPath + kSep + kFileName
kURL = os.path.expanduser('~/tmp/Logs/test.xlsx')

disc_catalog = {}
job_list = []
simulate = False
log_buffer = []

class Job:
    def __init__(self, job_number):
        self.job_number = job_number
        self.location = generate_job_url(job_number)
        self.is_archived = False
        self.size = get_size(self.location)
        self.on_server = True
        self.clean = False
        self.archive = None
        self.on_disc = None
        self.ignore = False
        self.inspect()

    def dump(self):
        shutil.rmtree(self.location)
        self.on_server = False

    def cleanup(self):
        # empty Trash folder
        try:
            shutil.rmtree('{0}/Trash'.format(self.location))
        except:
            pass

        # clean up Image Carrier folders
        """
        1. tokenize each file name in the image carrier folder, compile a list
        2. if a file has a single token, mark it to keep
        3. if a file has multiple tokens but none match the job number, mark to delete
        4. if a file has the job number token, mark it to keep
        5. if a file has a color token
            i. look for other files that have the same token
            ii. mark the most recent to keep
            iii. mark the remainder to delete
        6. if no files are marked to keep, mark all files to keep
        7. delete files marked for deletion

        """

        def tokenize(filename):
            strip_extension = filename.rpartition('.')[0]
            break_by_whitespace = strip_extension.split()
            break_by_underscore = []
            for this_section in break_by_whitespace:
                break_by_underscore = break_by_underscore + this_section.split('_')
            break_by_hyphen = []
            for this_section in break_by_underscore:
                break_by_hyphen = break_by_hyphen + this_section.split('-')
            tokens = break_by_hyphen
            return tokens

        # compile file list

        image_carrier_path = self.location + '/Deliverables/Image_Carriers'

        files = {}
        for this_tuple in os.walk(image_carrier_path):
            this_dir = this_tuple[0]
            if this_dir[0] != '.':  # ignore dot directories
                these_files = this_tuple[2]
                for this_file in these_files:
                    if this_file[0] != '.':  # ignore dot files and hidden stores
                        this_filepath = this_dir + '/' + this_file
                        files[this_filepath] = {}
                        files[this_filepath]['name'] = this_file
                        files[this_filepath]['modified'] = os.path.getmtime(this_filepath)
                        files[this_filepath]['tokens'] = tokenize(this_file)
                        files[this_filepath]['keep'] = True

        # clean out not image carriers from the deletion lists

        for this_filepath in files.keys():
            is_len = files[this_filepath]['name'].rsplit('.')[-1].lower() == 'len'
            is_tif = files[this_filepath]['name'].rsplit('.')[-1].lower() == 'tif'
            is_tiff = files[this_filepath]['name'].rsplit('.')[-1].lower() == 'tiff'
            if not(is_len or is_tif or is_tiff):
                try:
                    files.pop(this_filepath, None)
                except:
                    pass

        # check for a matching job number token, if not there don't keep

        for this_filepath in files.keys():
            if not (str(self.job_number) in files[this_filepath]['tokens']):
                files[this_filepath]['keep'] = False

        # accumulate color tokens

        color_tokens = {}
        for this_filepath in files.keys():
            if files[this_filepath]['keep']:
                color = files[this_filepath]['tokens'][-1]  # last token should be esko's ink color
                if not (color in color_tokens.keys()):  # color token not yet created
                    color_tokens[color] = []
                    color_tokens[color].append(this_filepath)
                else:
                    color_tokens[color].append(this_filepath)

        for this_color in color_tokens.keys():  # with each color
            if len(color_tokens[this_color]) > 1:  # if there are more than one file
                dates = []  # for sorting
                file_by_dates = {}  # reverse lookup to files
                for this_filepath in color_tokens[this_color]:  # for each filepath in a color
                    dates.append(files[this_filepath]['modified'])  # append the date to the sorting list
                    file_by_dates[
                        files[this_filepath]['modified']] = this_filepath  # add the filepath to the reverse lookup table
                dates.sort()  # sort the dates
                files[file_by_dates[dates.pop()]]['keep'] = True  # pop the highest from the sort stack and keep that file
                while dates:
                    files[file_by_dates[dates.pop()]]['keep'] = False  #set the others to drop

        # check for single token, do this last in case the other tests have marked the file for deletion

        for this_filepath in files.keys():
            if len(files[this_filepath]['tokens']) <= 2:
                files[this_filepath]['keep'] = True

        # using the keep field, dump the dead files

        for this_filepath in files.keys():
            if not (files[this_filepath]['keep']):
                try:
                    os.remove(this_filepath)
                except:
                    pass

        self.clean = True

    def inspect(self):
        my_walk = os.walk(self.location)
        path, dirs, files = my_walk.next()

        try:
            files.remove('.DS_Store')  # kill the Extreme IP mac share data
        except:
            pass

        try:
            dirs.remove('config')  # kill the hidden config dirs
        except:
            pass

        if len(dirs) == 1:  #could mean a disc folder already exists
            try:
                disc = int(dirs[0].lower().strip('disk#'))  #try to extract a disk number from the single directory
            except:
                raise Exception('Disk Folder not recognized.')  #no disc folder means a human needs to check the directory manually

            # if we get here, a disc got recognized and we need to make sure it goes into the excel file
            self.is_archived = True
            self.on_disc = disc

        elif len(dirs) == 0:
            try:
                raise Exception('Job Folder Is Empty.')  #no directories means something is wrong, get a human
            except:
                self.ignore = True
        else:
            self.is_archived = False
            self.ignore = False


class File:
    def __init__(self, url):
        self.location = url
        print(repr(url))
        self.name = url.rsplit('/',1)[1]
        self.size = get_size(url)
        self.is_placed = False
        self.in_disc = False

    def add2disc(self, disc):
        if self.size + disc.size >= kFullSize:
            return False
        else:
            shutil.move(self.location, disc.location)
            self.is_placed = True
            self.in_disc = disc
            disc.size += self.size
            if disc.size >= kFullSize:
                disc.is_full = True
            self.location = disc.location + kSep + self.name
            return True

class Archive:
    def __init__(self, job):
        self.job = job
        self.files = []

        # archive the job
        command = 'ditto'
        args = '-c -k --sequesterRsrc --keepParent'
        zip_name = str(job.job_number) + '.zip'
        zip_path = kWorkingPath + kSep + zip_name
        do_this = '{0} {1} {2} {3}'.format(command, args, job.location, zip_path)
        subprocess.call(do_this, shell=True)
        zip_size = get_size(zip_path)

        # break it up if required
        if zip_size >= kFullSize:
            command = 'zip'
            args = '-s 2g'
            split_path = kWorkingPath + kSep + '{0}_split.zip'.format(job.job_number)
            do_this = '{0} {1} {2} {3}'.format(command, args, split_path, zip_path)
            subprocess.call(do_this, shell=True)
            os.remove(zip_path)
            these_files = os.listdir(kWorkingPath)
            for this_file in these_files:
                if this_file.find('_split') > 0:  # contains the split file name
                    self.files.append(File(kWorkingPath + kSep + this_file))
        else:
            self.files.append(File(zip_path))
        job.is_archived = True
        job.archive = self


class Disc:
    def __init__(self, disc_number):
        self.disc_number = disc_number
        self.folder_name = kDiscFolderPrefix + str(self.disc_number).zfill(4)
        self.location = kBaseDisksPath + kSep + self.folder_name
        try:
            os.makedirs(self.location)
        except:
            pass
        self.size = get_size(self.location)
        if self.size >= kFullSize:
            self.is_full = True
        else:
            self.is_full = False
        self.contents = os.listdir(self.location)

class Manager:
    def __init__(self, url):
        self.job_list = []
        self.disc_catalog = []
        self.get_job_list(url)
        self.setup_disc_catalog()

    def get_job_list(self, url):
        """
        Setups up the master job list and attaches the Excel file log to the manager object.
        """
        try:
            self.wb = openpyxl.load_workbook(url)
        except:
            print('No excel file at {0)'.format(url))
            return False
        ws = self.wb.active
        max_row = int(ws.max_row) + 1
        max_col = ws.max_column

        r, c, c2 = 1, 1, 5

        while r < max_row:
            try:
                job_number = int(ws.cell(row=r, column=c).value)  # first entry is a job number
            except:
                r += 1
                continue
            if not (ws.cell(row=r, column=c2).value):  # disc already attached to job
                parsed_jobs = []
                # check for double entries
                for this_job in self.job_list:
                    parsed_jobs.append(this_job.job_number)
                if parsed_jobs.count(job_number) == 0:
                    self.job_list.append(Job(job_number))
            r += 1
        return True

    def setup_disc_catalog(self, url=kBaseDisksPath):
        """
        searches the base disc directory for existing disc folders, sizes them and populates the disc catalog
        """

        """
        1. get the list of folders in the staging directory, don't read deeper than the root level
        2. remove any non-disc folders from that listing
        3. with the remaining folders, calculate the size of each disc, store the value
        4. if no disc folders exists, prompt for a disc number to start with, create that folder, set it's size to 0
        """

        directories = os.walk(url).next()[1]

        for this_dir in directories:  # drop the non-disc dirs
            if this_dir.find('Disc') == -1:
                directories.remove(this_dir)

        if len(directories) > 0:  # one or more disc directories

            for index in range(len(directories)):  # drop the "Disc" prefix and convert the ids to ints
                directories[index] = int(directories[index].strip('Disc'))

            for this_disc_number in directories:
                self.disc_catalog.append(Disc(this_disc_number))
        else:  # no disc directories
            while not ('disc_number' in locals()):  # ask for a disc number until you get a valid one
                try:
                    disc_number = int(input('Enter a starting Disc #: '))
                except:
                    print('Disc number is invalid, please enter an integer disc number.')
            self.disc_catalog.append(Disc(disc_number))

        return True

    def get_last_disc(self):
        disc_numbers = []
        for this_disc in self.disc_catalog:
            disc_numbers.append(this_disc.disc_number)
        disc_numbers.sort()
        return disc_numbers[-1]



def dump_log(filename='log_' + str(int(time.time()))):
    '''

    :param filename: name of the log file to open
    :return: file object for the log file
    '''
    if not simulate:
        path = kWorkingPath
    else:
        path = kPath + '/tmp'
    url = path + kSep + filename
    log = open(url, 'w+')
    log.writelines(log_buffer)
    log.close()
    return url


def get_list(url):
    """

    :param url: The location of the excel file we're using
    :return: a list of dictionaries, containing 'id', 'row' and 'col' of the job in the excel file
    """

    try:
        wb = openpyxl.load_workbook(url)
    except:
        print('No excel file at {0)'.format(url))
    ws = wb.active
    max_row = int(ws.max_row) + 1
    max_col = ws.max_column

    r, c, c2 = 1, 1, 5

    while r < max_row:
        item = {}
        try:
            item['id'] = int(ws.cell(row=r, column=c).value)
            item['row'] = r
            item['col'] = c
        except:
            r += 1
            continue
        if not (ws.cell(row=r, column=c2).value):
            job_list.append(item)
        r += 1
    return 0


def generate_job_url(job):
    """

    :param job: the
     job number to generate the URL for
    :return: the string of the URL to the job on the server
    """

    if type(job) != type(''):
        job = str(job)
    folderNumber = job[len(job) - 1:len(job)]  # get the last number of the job
    return kBaseJobsPath + kJobFolderPrefix + folderNumber + kSep + job


def generate_working_url(job):
    """

    :param job: the job number to generate a working folder for
    :return: the string of the URL of that working folder
    """
    if type(job) != type(''):
        job = str(job)
    return kWorkingPath + kSep + job


def generate_disc_url(disc):
    """

    :param disc: the disc number to generate path to string for
    :return: the path to string
    """
    if not isinstance(disc, str):
        disc = str(disc).zfill(4)
    return kBaseDisksPath + kDiscFolderPrefix + disc


def dump_trash(job):
    """

    :param job: The job number to dump the trash of
    :return:
    """
    deleteMe = generate_job_url(job) + kTrashFolderPrefix
    if not simulate:
        try:
            shutil.rmtree(deleteMe)
        except:
            pass
    else:
        print(deleteMe)
    return 0


def clean_IC(job):
    """

    :param job: job number to clean up the image carrier folder for
    :return: 0 if successful
    """

    """
    1. tokenize each file name in the image carrier folder, compile a list
    2. if a file has a single token, mark it to keep
    3. if a file has multiple tokens but none match the job number, mark to delete
    4. if a file has the job number token, mark it to keep
    5. if a file has a color token
        i. look for other files that have the same token
        ii. mark the most recent to keep
        iii. mark the remainder to delete
    6. if no files are marked to keep, mark all files to keep
    7. delete files marked for deletion

    """

    def tokenize(filename):
        strip_extension = filename.rpartition('.')[0]
        break_by_whitespace = strip_extension.split()
        break_by_underscore = []
        for this_section in break_by_whitespace:
            break_by_underscore = break_by_underscore + this_section.split('_')
        break_by_hyphen = []
        for this_section in break_by_underscore:
            break_by_hyphen = break_by_hyphen + this_section.split('-')
        tokens = break_by_hyphen
        return tokens

    # compile file list

    image_carrier_path = generate_job_url(job) + '/Deliverables/Image_Carriers'

    files = {}
    for this_tuple in os.walk(image_carrier_path):
        this_dir = this_tuple[0]
        if this_dir[0] != '.':  # ignore dot directories
            these_files = this_tuple[2]
            for this_file in these_files:
                if this_file[0] != '.':  # ignore dot files and hidden stores
                    this_filepath = this_dir + '/' + this_file
                    files[this_filepath] = {}
                    files[this_filepath]['name'] = this_file
                    files[this_filepath]['modified'] = os.path.getmtime(this_filepath)
                    files[this_filepath]['tokens'] = tokenize(this_file)
                    files[this_filepath]['keep'] = True

    # check for a matching job number token, if not there don't keep

    for this_filepath in files.keys():
        if not (str(job) in files[this_filepath]['tokens']):
            files[this_filepath]['keep'] = False

    # accumulate color tokens

    color_tokens = {}
    for this_filepath in files.keys():
        if files[this_filepath]['keep']:
            color = files[this_filepath]['tokens'][-1]  # last token should be esko's ink color
            if color_tokens.get(color, True):  # color token not yet created
                color_tokens[color] = []
                color_tokens[color].append(this_filepath)
            else:
                color_tokens[color].append(this_filepath)

    for this_color in color_tokens.keys():  # with each color
        if len(color_tokens[this_color]) > 1:  # if there are more than one file
            dates = []  # for sorting
            file_by_dates = {}  # reverse lookup to files
            for this_filepath in color_tokens[this_color]:  # for each filepath in a color
                dates.append(files[this_filepath]['modified'])  # append the date to the sorting list
                file_by_dates[
                    files[this_filepath]['modified']] = this_filepath  # add the filepath to the reverse lookup table
            dates.sort()  # sort the dates
            files[file_by_dates[dates.pop()]]['keep'] = True  # pop the highest from the sort stack and keep that file
            while dates:
                files[file_by_dates[dates.pop()]]['keep'] = False  #set the others to drop

    # check for single token, do this last in case the other tests have marked the file for deletion

    for this_filepath in files.keys():
        if len(files[this_filepath]['tokens']) == 1:
            files[this_filepath]['keep'] = True

    # using the keep field, dump the dead files

    for this_filepath in files.keys():
        if not (files[this_filepath]['keep']):
            if not simulate:
                try:
                    os.remove(this_filepath)
                    log_buffer.append('Deleted: {0}\n'.format(this_filepath))
                except:
                    log_buffer.append(('Could not delete: {0}\n'.format(this_filepath)))
            else:
                log_buffer.append('Will delete: {0}\n'.format(this_filepath))

    '''
    for this_filepath in files.keys():
        print('path: {0}, settings: {1}'.format(this_filepath, repr(files[this_filepath])))
    '''

    return 0


def copy_job(job):
    """

    :param job: The job number to copy to the work folder
    :return: the copy's url
    """
    copyMeFrom = generate_job_url(job)
    copyMeTo = generate_working_url(job)
    command = 'ditto'
    do_this = [command, copyMeFrom, copyMeTo]
    if not simulate:
        try:
            subprocess.call(do_this)
            log_buffer.append("Copied: {0} to {1}\n".format(copyMeFrom, copyMeTo))
        except:
            log_buffer.append('Unable to copy: {0} to {1}\n'.format(copyMeFrom, copyMeTo))
    else:
        log_buffer.append('Will copy: {0} to {1}\n'.format(copyMeFrom, copyMeTo))
    return copyMeTo


def zip_job(job):
    """

    :param job: The job number to zip using Archive Utility
    :return: zip file path
    """

    command = 'ditto'
    args = '-c -k --sequesterRsrc --keepParent'
    working_path = generate_job_url(job)
    zip_path = kWorkingPath + '/' + str(job) + '.zip'
    do_this = '{0} {1} {2} {3}'.format(command, args, working_path, zip_path)

    if not simulate:
        try:
            subprocess.call(do_this, shell=True)
            log_buffer.append('Zipped: {0} to {1}\n'.format(job, zip_path))
        except:
            log_buffer.append('Unable to zip: {0} to {1}\n'.format(job, zip_path))
    else:
        log_buffer.append('Will zip: {0} to {1}\n'.format(job, zip_path))
    return zip_path


def inspect_job(job):
    """

    :param job: job number to inspect
    :return: true, job needs to be archived or false, job has issues
    """

    path = generate_job_url(job)
    my_walk = os.walk(path)
    path, dirs, files = my_walk.next()

    try:
        files.remove('.DS_Store')  # kill the Extreme IP mac share data
    except:
        pass

    try:
        dirs.remove('config')  # kill the hidden config dirs
    except:
        pass

    try:
        if len(files) > 0:
            raise Exception('Path Not Empty.')  #loose files in the directory means something is wrong, get a human

        if len(dirs) > 1:  #more than one directory means there's something to archive
            return True
        if len(dirs) == 1:  #could mean a disc folder already exists
            try:
                disc = int(dirs[0].lower().strip('disk#'))  #try to extract a disk number from the single directory
            except:
                raise Exception('Disk Folder not recognized.')  #no disc folder means a human needs to check the directory manually

            # if we get here, a disc got recognized and we need to make sure it goes into the excel file
            for this_item in job_list:
                if this_item['id'] == job:
                    this_item['disc'] = disc
            return False  # but the job doesn't need to be archived
        if len(dirs) == 0:
            raise Exception('Disk Folder Is Empty.')  #no directories means something is wrong, get a human
    except:
        # if any exception gets raised on inspection, we can't continue with archiving the job
        for a in job_list:
            if a['id'] == job:
                job_list.remove(a)  #remove the job from the master list
        return False


def bucket_job(zip_url):
    """

    :param zip: zip to put into disc folder
    :return: disc number ZIP was put in
    """
    """
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
    zip_size = get_size(zip_url)
    zip_name = zip_url.rsplit('/', 1)[-1]

    discs = disc_catalog.keys()  # get the disc numbers
    discs.sort()  # sort the disc numbers from low to high

    zip_placed = False  # zip has not yet been placed

    bucket = 0

    for this_disc in discs:
        if zip_size + disc_catalog[this_disc]['size'] >= kFullSize:
            pass  # next disc
        else:
            zip_placed = True
            if not simulate:
                # move the file to the directory
                shutil.move(zip_url, disc_catalog[this_disc]['path']+kSep+zip_name)
                log_buffer.append('{0} moved to disc {1}, @ {2}\n'.format(zip_name, this_disc, disc_catalog[this_disc]['path']))
                # update the size in the catalog
                disc_catalog[this_disc]['size'] += zip_size
                bucket = this_disc
            else:
                log_buffer.append('{0} will be moved to disc {1}, @ {2}\n'.format(zip_name, this_disc, disc_catalog[this_disc]['path']))

    if not zip_placed:  # zip didn't get placed in the disc folders
        new_disc = create_disc()
        if not simulate:
            # move the zip file into the new disc
            shutil.move(zip_url, disc_catalog[new_disc]['path']+kSep+zip_name)
            log_buffer.append('{0} moved to disc {1}, @ {2}\n'.format(zip_name, new_disc, disc_catalog[new_disc]['path']))
            # update the new disc's size
            disc_catalog[new_disc]['size'] += zip_size
            bucket = new_disc
        else:
            log_buffer.append('{0} moved to disc {1}, @ {2}\n'.format(zip_name, new_disc, disc_catalog[new_disc]['path']))

    return bucket


def inspect_discs():
    """
    searches the staging directory for existing disc folders, sizes them and populates the disc catalog
    :return:
    """

    """
    1. get the list of folders in the staging directory, don't read deeper than the root level
    2. remove any non-disc folders from that listing
    3. with the remaining folders, calculate the size of each disc, store the value
    4. if no disc folders exists, prompt for a disc number to start with, create that folder, set it's size to 0
    """
    staging = kBaseDisksPath

    directories = os.walk(staging).next()[1]

    for this_dir in directories:  # drop the non-disc dirs
        if this_dir.find('Disc') == -1:
            directories.remove(this_dir)

    if len(directories) > 0:  # one or more disc directories

        for index in range(len(directories)):  # drop the "Disc" prefix and convert the ids to ints
            directories[index] = int(directories[index].strip('Disc'))

        for this_dir in directories:
            disc_catalog[this_dir] = {}
            disc_catalog[this_dir]['path'] = staging + kDiscFolderPrefix + str(this_dir).zfill(4)
            disc_catalog[this_dir]['size'] = get_size(disc_catalog[this_dir]['path'])
    else:  # no disc directories
        while not ('disc' in locals()):  # ask for a disc number until you get a valid one
            try:
                disc = int(input('Enter a starting Disc #: '))
            except:
                print('Disc number is invalid, please enter an integer disc number.')

        os.mkdir(staging + kDiscFolderPrefix + str(disc).zfill(4))  # create a new disc directory
        log_buffer.append('Created Directory: {0}\n'.format(staging + kDiscFolderPrefix + str(disc).zfill(4)))
        disc_catalog[disc] = {}  # add an entry to the catalog
        disc_catalog[disc]['path'] = staging + kDiscFolderPrefix + str(disc).zfill(4)
        disc_catalog[disc]['size'] = 0

    return disc_catalog


def get_size(path):
    total_size = 0
    try:
        os.walk(path).next()
        for dir_path, dir_names, filenames in os.walk(path):
            for f in filenames:
                fp = os.path.join(dir_path, f)
                total_size += os.path.getsize(fp)
    except:
        total_size = os.path.getsize(path)
    return total_size


def check_disc(disc):
    """
    :param disc: the disc number of the folder to check for size
    :return: full or not
    """

    try:
        if disc_catalog[disc]['size'] >= kFullSize:
            print('Disc {0} is full.'.format(disc))
            return True
        else:
            print('Disc {0} has room.'.format(disc))
            return False
    except KeyError:
        print('Disc {0} Does Not Exist.'.format(disc))
        return -1


def create_disc():
    """
    :return the number of the new disc
    """
    discs = disc_catalog.keys()
    discs.sort()
    last_disc = discs.pop()
    new_disc = last_disc + 1
    if not simulate:
        try:
            os.mkdir(generate_disc_url(new_disc))
            disc_catalog[new_disc]['path'] = generate_disc_url(new_disc)
            disc_catalog[new_disc]['size'] = 0
            log_buffer.append('Created new disc folder: {0} at {1}\n'.format(new_disc, generate_disc_url(new_disc)))
        except:
            log_buffer.append(
                'Unable to create new disc folder: {0} at {1}\n'.format(new_disc, generate_disc_url(new_disc)))
    else:
        log_buffer.append('Will create new disc folder: {0} at {1}\n'.format(new_disc, generate_disc_url(new_disc)))
    return new_disc


def dump_job(job):
    """

    :param job: the job number to dump the contents on
    :return:
    """
    if not simulate:
        try:
            shutil.rmtree(generate_job_url(job))
            log_buffer.append('Dumped job#: {0} from {1}\n'.format(job, generate_job_url(job)))
        except:
            log_buffer.append('Unable to dump job#: {0} from {1}\n'.format(job, generate_job_url(job)))
    else:
        log_buffer.append('Will dump job#: {0} from {1}\n'.format(job, generate_job_url(job)))
    return 0


def tag_job(job, disc):
    """

    :param job: the job number to add the disc tag folder to
    :param disc: the disc number for the tag
    :return:
    """
    if not simulate:
        try:
            os.makedirs(generate_job_url(job) + kSep + kDiscFolderPrefix + str(disc).zfill(4))
            log_buffer.append('Tagged Job# {0} with Disc# {1}\n'.format(job, disc))
        except:
            log_buffer.append('Unable to tag Job# {0} with Disc# {1}\n'.format(job, disc))
    else:
        log_buffer.append('Will tag job# {0} with Disc# {1}\n'.format(job, disc))
    return 0


def compile_excel_tags(job, disc):
    '''

    :param job_list: the list of dictionaries gathered by get_list
    :param job: a job number
    :param disc:  a disc number
    :return: the modified listing
    '''

    for this_item in job_list:
        if this_item['id'] == job:
            this_item['disc'] = disc
    return 0


def dump_list_to_excel(url):
    '''

    :param url: path of the excel file to update
    :return:
    '''
    wb = openpyxl.load_workbook(url)
    ws = wb.active
    for this_item in job_list:
        if not simulate:
            try:
                ws.cell(row=this_item['row'], column=4).value = this_item['disc']
                log_buffer.append(
                    'Updated Excel File {0}, job# {1} to disc# {2}\n'.format(url, this_item['id'], this_item['disc']))
            except:
                log_buffer.append(
                    'Problem updating Excel File {0}, job# {1} to disc# {2}\n'.format(url, this_item['id'],
                                                                                      this_item['disc']))
    wb.save(url)
    return 0


if __name__ == "__main__":
    mngr = Manager(kURL)
    for this_job in mngr.job_list:
        print('Job: {0}, @:{1}\n'.format(this_job.job_number, this_job.location))
    for this_disc in mngr.disc_catalog:
        print('Disc: {0}, @:{1}\n'.format(this_disc.disc_number, this_disc.location))

    # code for bucketing the jobs
    for this_job in mngr.job_list:
        if not this_job.ignore and not this_job.is_archived:
            this_job.cleanup()
            this_job.archive = Archive(this_job)
            for this_file in this_job.archive.files:
                index = 0
                while index < len(mngr.disc_catalog):
                    try:
                        if not(this_file.add2disc(mngr.disc_catalog[index])):
                            index += 1
                        else:
                            break
                    except:
                        pass
                if not this_file.is_placed:
                    new_disc_number = mngr.get_last_disc() + 1
                    new_disc = Disc(new_disc_number)
                    mngr.disc_catalog.append(new_disc)
                    this_file.add2disc(new_disc)




    """
    get_list(kURL)
    # print(repr(job_list))
    # cleanup the list
    for this_job in job_list:
        inspect_job(this_job['id'])
    print(repr(job_list))
    # setup the disc catalog
    inspect_discs()
    # for each job
    for this_job in job_list:
        if not this_job.has_key('disc'):
            dump_trash(this_job['id']) # empty the trash folder
            clean_IC(this_job['id']) # clean out the image carriers
            this_zip = zip_job(this_job['id'])
            print(this_zip)
            this_job['disc'] = bucket_job(this_zip) # zip and sort the the zip into a folder
            dump_job(this_job['id']) # empty the job folder
            tag_job(this_job['id'], this_job['disc']) # create a disc folder tag
        compile_excel_tags(this_job['id'], this_job['disc']) # compile the excel data
    dump_list_to_excel(kURL) # write the compiled data to the xl doc
    dump_log() # dump the log file
    """
