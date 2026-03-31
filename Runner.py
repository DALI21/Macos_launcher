from Lib_RuneurJenkins import (
    CreatFolder,
    CreatFolderAllure,
    ExcuteRobotTest,
    AddJenkinsLogToRobot,
    ExcutePythonTest,
    ExcuteNPMTest,
    ParserLOGPythonTest,
    generateJson,
)
import argparse


def parse_args():
    parser = argparse.ArgumentParser(
        prog="RunnerTest",
        description="Generic test runner for Robot, Python, or NPM tests",
    )

    parser.add_argument(
        "--tag",
        type=str,
        default=None,
        help="Tag used to filter or label the test execution",
    )

    parser.add_argument(
        "--test-file",
        required=True,
        type=str,
        help="Path to the test file (.robot, .py, or other for NPM test)",
    )

    parser.add_argument(
        "--test-exec-key",
        type=str,
        default=None,
        help="Execution key (e.g. Xray / test management execution id)",
    )

    parser.add_argument(
        "--arg",
        type=str,
        default="",
        help="Additional arguments passed to the test runner "
             "(e.g. Robot Framework --variable options)",
    )

    parser.add_argument(
        "--workspace",
        type=str,
        default="",
        help="Workspace path (e.g. Jenkins workspace root)",
    )

    return parser.parse_args()


def main() -> None:
    """Main entry point: parses args and dispatches to the correct runner."""
    args = parse_args()

    workspace = args.workspace
    tag = args.tag
    test_exec_key = args.test_exec_key
    test_file = args.test_file
    arg = args.arg

    # Create logs directory
    Logs_Directory = CreatFolder(test_file)
    Logs_Directory1 = ""

    if workspace:
        Logs_Directory1 = Logs_Directory
        print("path logs directory", Logs_Directory1)
    else:
        Logs_Directory1 = Logs_Directory

    # Run according to the test file type
    if test_file.endswith(".robot"):
        # Robot Framework test
        Dossier_Allure = CreatFolderAllure(Logs_Directory1)
        ExcuteRobotTest(Logs_Directory1, Dossier_Allure, test_file, arg, tag)
        AddJenkinsLogToRobot(Logs_Directory)

    elif test_file.endswith(".py"):
        # Python test
        output = ExcutePythonTest(Logs_Directory1, test_file, arg)
        status = ParserLOGPythonTest(output)
        generateJson(Logs_Directory, test_exec_key, tag, status)

    else:
        # Assume NPM or other type of test
        output = ExcuteNPMTest(test_file)
        print(output)
        status = ParserLOGPythonTest(output)
        generateJson(Logs_Directory, test_exec_key, tag, status)


if __name__ == "__main__":
    main()
