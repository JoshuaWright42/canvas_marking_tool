#!/usr/bin/env python3


from fileinput import close
from heapq import merge
import json
import os
import requests
from pygments import highlight
from pygments.lexers import get_lexer_for_filename
from pygments.formatters import get_formatter_by_name
import pdfkit
from canvasapi import Canvas
from PIL import Image
from docx2pdf import convert
from PyPDF2 import PdfMerger



CONFIG_FILE = "config.json"
CONFIG_DATA = json.load(open(CONFIG_FILE))
WORKING_DIR = "temp/"
DOWNLOAD_DIR = "downloads/"
ERROR = "error"


def load_course():
    '''Load the configuration from file and create a session connecting to the relevant course.

    Returns the course object.
    '''

    api_key = CONFIG_DATA["API_KEY"]

    # use environment variable api key if present -- takes precedent over config file
    try:
        api_key = os.environ["CANVAS_KEY"]
    except:
        pass

    # connect to canvas and retrieve course
    canvas = Canvas(CONFIG_DATA["API_URL"], api_key)
    course = canvas.get_course(CONFIG_DATA["COURSE_ID"])
    return course


def get_all_valid_submissions(course):
    '''Find all valid submissions for all tasks.

    Returns a dictionary of {assignment: list of submissions}.

    For a submission to be valid it must be submitted before the task due date, and
    not have been graded yet.
    '''
    
    valid_submissions = {}
    for a in course.get_assignments():
        print("- pulling submissions for", a.name)
        # response = requests.get(a.submissions_download_url)
        # subfile = open("submissions" + str(a.id) + ".zip", "wb").write(response.content)
        # close()
        # wget.download(a.submissions_download_url)
        submissions = [s for s in a.get_submissions() 
            if s.submitted_at is not None # something was submitted
            and s.submitted_at_date <= a.due_at_date # it's not overdue
            and (s.graded_at is None or s.graded_at_date < s.submitted_at_date)] # it hasn't ever been graded, or the grade there is for an older submission
        if len(submissions) > 0:
            valid_submissions[a] = submissions
            for s in submissions:
                user = course.get_user(s.user_id)
                compile_pdf(s, user, a.name)
                # for a in s.attachments:
                #     print(a['filename'])
                #     response = requests.get(a['url'])
                #     subfile = open(a['display_name'], "wb").write(response.content)
                #     close()
    
    return valid_submissions


def compile_pdf(submission, student, task_name):
    path = WORKING_DIR + student.login_id + "/"
    if not (os.path.exists(path)):
        os.makedirs(path)
    merger = PdfMerger()
    for a in submission.attachments:
        filename = ""
        if validate_file_type(a["display_name"]):
            filename = convert_to_pdf(a, path)
        else:
            print("Invalid filetype: " + a["display_name"])
            filename = ERROR
        if (filename != ERROR):
            merger.append(filename)
    if not (os.path.exists(DOWNLOAD_DIR)):
        os.makedirs(DOWNLOAD_DIR)
    merger.write(DOWNLOAD_DIR + task_name + "_" + student.login_id + "_" + student.name + ".pdf")
    merger.close

def get_extension(filename):
    name_split = str.split(filename, ".")
    extension = name_split[len(name_split) - 1].lower()
    return extension


def validate_file_type(filename):
    extension = get_extension(filename)
    for e in CONFIG_DATA["VALID_EXTENSTIONS"]:
        if extension == e: return True
    return False


def convert_to_pdf(attachment, path):
    response = requests.get(attachment["url"])
    filename = path + attachment["display_name"]
    open(filename, "wb").write(response.content)
    close()
    extension = get_extension(attachment["display_name"])
    match extension:
        case "cs":
            filename = code_to_pdf(filename)
        case "png" | "jpeg" | "jpg" | "bmp":
            filename = img_to_pdf(filename)
        case "docx":
            filename = docx_to_pdf(filename)
        case "pdf":
            pass
        case _:
            print("Failed to convert unrecognised extension: " + extension)
            filename = ERROR
    return filename


def code_to_pdf(filename):
    with (open(filename, "r")) as f:
        html = open(filename + ".html", "w")
        code = ""
        for line in f:
            code += line
        html.write(highlight(code, get_lexer_for_filename(filename), get_formatter_by_name("html", full=True, linenos=True, nobackground=True, title=filename)))
        close()
    pdfkit.from_file(filename + ".html", filename + ".pdf")
    return filename + ".pdf"


def img_to_pdf(filename):
    image = Image.open(r"" + filename)
    pdf = image.convert("RGB")
    pdf.save(r"" + filename + ".pdf")
    return filename + ".pdf"


def docx_to_pdf(filename):
    convert(filename, filename + ".pdf")
    return filename + ".pdf"



def speedgrader_link(course, assignment, submission):
    '''Get a speedgrader formatted link for viewing a specific submission on an assignment.'''

    return "https://swinburne.instructure.com/courses/" + str(course.id) + "/gradebook/speed_grader?assignment_id=" + str(assignment.id) + "&student_id=" + str(submission.user_id)


def find_lab(course, student_user_id):
    '''Get string representation for the lab a specific student in a course is enrolled in.'''

    course_enrollments = [e for e in course.get_enrollments(user_id=student_user_id)]
    for en in course_enrollments:
        section = course.get_section(en.course_section_id)
        if "Lab" in section.name:
            return section.name

    return "No lab"


def generate_csv(course, valid_submissions):
    '''Generate a csv called "marking.csv containing a summary of tasks that need feedback,
    and links to the speedgrader to show each task.
    '''

    print("Generating csv...")
    student_details = {} # cache student details to avoid unnecessary requests; {userid: ("name, student_id, lab")}

    with open("marking.csv", "w+") as f:
        f.write("Task, Student, Student ID, Lab, Submitted on, Times assessed, Link\n")
        for assignment, submissions in valid_submissions.items():
            print("- adding entries for", assignment.name)
            for s in submissions:
                student_string = student_details.get(s.user_id)
                if not student_string:
                    user = course.get_user(s.user_id)
                    student_string = user.name + "," + str(user.login_id) + "," + find_lab(course, s.user_id)
                    student_details[s.user_id] = student_string
                f.write(assignment.name + ","
                    + student_string + ","
                    + s.submitted_at_date.strftime("%d %B") + ","
                    + str(s.attempt) + ","
                    + speedgrader_link(course, assignment, s)
                    + "\n")


if __name__ == "__main__":
    course = load_course()
    print("Retrieving submissions for", course.name, "needing feedback...")
    valid_submissions = get_all_valid_submissions(course)
    generate_csv(course, valid_submissions)