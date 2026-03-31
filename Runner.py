import os
import shutil
import subprocess
import datetime
import sys
import xml.etree.ElementTree as ET
import json
import platform

# Windows-only imports for UFT (guarded later)
try:
    import win32com, win32com.client  # type: ignore
except ImportError:
    win32com = None

from lxml import etree


# ============================================================
# OS-agnostic base paths
# ============================================================

system = platform.system()

# Locate repo root dynamically from this file:
# Assumes this file is at: <ROOT>/source/runners/Lib_RuneurJenkins.py
CURRENT_FILE = os.path.abspath(__file__)
RUNNERS_DIR = os.path.dirname(CURRENT_FILE)      # .../source/runners
SOURCE_DIR = os.path.dirname(RUNNERS_DIR)        # .../source
BASE_ROOT = os.path.dirname(SOURCE_DIR)          # .../Cubestudio_AutoValid

# Temp directory per OS
if system == "Windows":
    base_temp = r"C:\Temp"
else:
    base_temp = "/tmp"

# Common layout under BASE_ROOT for all OSes
ROBOT_WORK_DIR = os.path.join(BASE_ROOT, "source", "Main_scenarios", "RFW")
PYTHON_WORK_DIR = os.path.join(BASE_ROOT, "source", "Main_scenarios", "Python")
NPM_WORK_DIR = os.path.join(BASE_ROOT, "source", "utilities", "CoopAPI", "CoopApi")


# ============================================================
# Helper: Run without Jenkins
# ============================================================

def RunTestWithoutJenkins(test_file):
    """
    Create folder with name of test + date + time for non-Jenkins runs.
    Returns Logs_Directory path.
    """

    if test_file.endswith(".robot"):
        test_file_no_ext = test_file[:-6]
    elif test_file.endswith(".py"):
        test_file_no_ext = test_file[:-3]
    else:
        test_file_no_ext = test_file

    parts = test_file_no_ext.replace("\\", "/").split("/")
    Name_test_file = parts[-1].split(".")[0]

    job_name = os.environ.get("JOB_NAME")
    if job_name is None:
        formatted_time = datetime.datetime.now().strftime("%Y-%m-%d_%H_%M_%S")
        job_name = f"{Name_test_file}_{formatted_time}"

    Logs_Directory = os.path.join(base_temp, job_name, Name_test_file)
    os.makedirs(Logs_Directory, exist_ok=True)

    return Logs_Directory


# ============================================================
# Workspace / Allure folders
# ============================================================

def CreatFolder(test_file):
    """
    Create workspace folder for Jenkins; if not under Jenkins, fallback to RunTestWithoutJenkins.
    Returns Logs_Directory.
    """
    job_name = os.environ.get("JOB_NAME")
    build_number = os.environ.get("BUILD_NUMBER")

    if job_name is not None and build_number is not None:
        Logs_Directory = os.path.join(base_temp, job_name, build_number)
        os.makedirs(Logs_Directory, exist_ok=True)
    else:
        print("JOB_NAME and BUILD_NUMBER do not exist. Running without Jenkins setup.")
        Logs_Directory = RunTestWithoutJenkins(test_file)

    return Logs_Directory


def CreatFolderAllure(Logs_Directory):
    """
    Create Allure results folder inside Logs_Directory.
    Returns Dossier_Allure path.
    """
    os.makedirs(Logs_Directory, exist_ok=True)

    Dossier_Allure = os.path.join(Logs_Directory, "allure-results")

    # Remove existing folder if any
    if os.path.exists(Dossier_Allure):
        shutil.rmtree(Dossier_Allure)

    os.mkdir(Dossier_Allure)
    return Dossier_Allure


# ============================================================
# Test execution (Robot, Python, NPM)
# ============================================================

def ExcuteRobotTest(Logs_Directory, Dossier_Allure, Robot_Test, arg, tag):
    """
    Execute Robot Framework test.
    Inputs:
        Logs_Directory: workspace logs path
        Dossier_Allure: Allure results path
        Robot_Test: path to .robot test file (absolute or relative)
        arg: extra robot CLI arguments string
        tag: test tag (Xray, RF include, etc.)
    """

    # Option A: run from common RFW root (repo layout-based)
    if os.path.isdir(ROBOT_WORK_DIR):
        os.chdir(ROBOT_WORK_DIR)
    else:
        # Fallback: directory of the test file
        os.chdir(os.path.dirname(os.path.abspath(Robot_Test)))

    # Build command: use same Python as runner, call robot as module
    command = [
        sys.executable,
        "-m", "robot",
        "--variable", f"Workspace:{Logs_Directory}",
        "--listener", f"allure_robotframework;{Dossier_Allure}",
        "-d", Logs_Directory,
    ]

    if tag:
        command += ["--variable", f"tag:{tag}"]

    if arg:
        # simple split of extra args string
        command += arg.split()

    # Robot test file (can be absolute or relative)
    command.append(Robot_Test)

    print("Executing Robot command:", " ".join(command))
    returncode = subprocess.call(command)

    if returncode != 0:
        print(f"Robot tests failed with exit code {returncode}")

    return returncode


def ExcutePythonTest(Logs_Directory, python_Test, arg):
    """
    Execute Python test.
    Inputs:
        Logs_Directory: workspace logs path
        python_Test: path to Python test (absolute or relative)
        arg: extra arguments string
    Returns:
        output: combined stdout of the Python test
    """

    if os.path.isdir(PYTHON_WORK_DIR):
        os.chdir(PYTHON_WORK_DIR)
    else:
        os.chdir(os.path.dirname(os.path.abspath(python_Test)))

    command = f'"{sys.executable}" "{python_Test}" "{Logs_Directory}" {arg}'
    print("Executing Python test:", command)
    output = subprocess.check_output(command, shell=True, universal_newlines=True)
    print(output)
    return output


def ExcuteNPMTest(Test):
    """
    Execute NPM test command.
    Input:
        Test: NPM script suffix (as used in 'npm run start:<Test>')
    Returns:
        output: combined stdout/stderr text
    """
    env = os.environ.copy()

    if os.path.isdir(NPM_WORK_DIR):
        os.chdir(NPM_WORK_DIR)
    else:
        os.chdir(os.getcwd())

    command = f"npm run start:{Test}"
    print("Executing NPM test:", command)

    p = subprocess.Popen(
        command,
        shell=True,
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
        text=True,
        env=env,
    )
    output, errors = p.communicate()
    if errors:
        output = output + "\n" + f"[FAIL]: {errors}"
    return output


def ParserLOGPythonTest(output):
    """
    Parse output text from Python or NPM tests and return PASS/FAIL.
    """
    if "FAIL" in output:
        return "FAIL"
    else:
        return "PASS"


# ============================================================
# UFT (Windows-only)
# ============================================================

def ExcuteUFTTest(testfile, logsDirectory):
    """
    Run an UFT test (Windows only). Returns PASS or FAIL.
    """
    if system != "Windows" or win32com is None:
        raise RuntimeError("ExcuteUFTTest is only supported on Windows with pywin32 installed.")

    qtp = win32com.client.Dispatch("QuickTest.Application")
    qtp.Launch()
    qtp.Visible = False

    qtp.Open(testfile)

    qtResultsOpt = win32com.client.Dispatch("QuickTest.RunResultsOptions")

    qtp.Test.Settings.Run.IterationMode = "rngIterations"
    qtp.Test.Settings.Run.OnError = "NextStep"

    qtAutoExportResultsOpts = qtp.Options.Run.AutoExportReportConfig
    qtResultsOpt.ResultsLocation = logsDirectory

    qtAutoExportResultsOpts.AutoExportResults = True
    qtAutoExportResultsOpts.StepDetailsReport = True
    qtAutoExportResultsOpts.DataTableReport = True
    qtAutoExportResultsOpts.LogTrackingReport = True
    qtAutoExportResultsOpts.ScreenRecorderReport = True
    qtAutoExportResultsOpts.SystemMonitorReport = False

    qtp.Test.Run(qtResultsOpt)
    results = qtp.Test.LastRunResults

    print("Test Results:")
    print("Status: ", results.Status)

    if results.Status == "Failed":
        return "FAIL"
    elif results.Status == "Passed":
        return "PASS"
    else:
        return "FAIL"


# ============================================================
# Reporting: XML / JSON / Jenkins info
# ============================================================

def generateXML(status, tag, directory, testName):
    """
    Generate TestNG-like XML for launched test.
    """

    if status == "FAIL":
        root = etree.Element("testng-results", skipped="0", failed="1", ignored="0", total="1", passed="0")
    else:
        root = etree.Element("testng-results", skipped="0", failed="0", ignored="0", total="1", passed="1")

    suite = etree.SubElement(root, "suite")
    test = etree.SubElement(suite, "test", name=testName)
    test_class = etree.SubElement(test, "class", name="MyTestClass")
    test_method = etree.SubElement(test_class, "test-method")
    test_method.set("status", status)

    attributes = etree.SubElement(test_method, "attributes")
    test_attribute = etree.SubElement(attributes, "attribute", name="test")
    test_attribute.text = tag

    xml_string = etree.tostring(root, encoding="utf-8", pretty_print=True, xml_declaration=True)
    filepath = os.path.join(directory, "output.xml")
    with open(filepath, "wb") as f:
        f.write(xml_string)


def AddJenkinsLogToRobot(logsDirectory):
    """
    Update output.xml of a robot test and add the Jenkins log (console URL or message).
    """
    filepath = os.path.join(logsDirectory, "output.xml")
    job_name = os.environ.get("JOB_NAME")

    if job_name is None:
        print("Running without Jenkins")
        console = "Job was executed without Jenkins"
    else:
        buildurl = os.environ.get("BUILD_URL")
        console = buildurl + "console"

    tree = ET.parse(filepath)
    root = tree.getroot()

    test_element = root.find(".//test")
    status_elements = test_element.findall("status")
    last_status_element = status_elements[-1]

    last_status_element.text = console
    tree.write(filepath)


def generateJson(logsDirectory, test_execution_key, test_key, status):
    """
    Generate output.json for the launched test (Xray format).
    """
    filepath = os.path.join(logsDirectory, "output.json")
    job_name = os.environ.get("JOB_NAME")

    if job_name is None:
        print("Running without Jenkins")
        console = "Job was executed without Jenkins"
    else:
        buildurl = os.environ.get("BUILD_URL")
        console = buildurl + "console"

    data = {
        "testExecutionKey": test_execution_key,
        "tests": [
            {
                "testKey": test_key,
                "status": status,
                "comment": console,
            }
        ],
    }

    with open(filepath, "w") as outfile:
        json.dump(data, outfile, indent=4)
