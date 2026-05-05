from Lib_RuneurJenkins import *
import sys
import platform
import argparse
import pydantic
import pydantic_argparse


class Arguments(pydantic.BaseModel):
    """Simple Command-Line Arguments."""
    tag: str = pydantic.Field(description="a required string", default=None)
    test_file: str = pydantic.Field(description="a required integer")
    test_exec_key: str = pydantic.Field(description="a required integer", default=None)
    arg: str = pydantic.Field(description="a required integer", default="")
    workspace: str = pydantic.Field(description="a required integer", default="")


def parse_args_macos():
    parser = argparse.ArgumentParser(
        prog="Example Program",
        description="Example Description",
    )
    parser.add_argument("--tag", required=True)
    parser.add_argument("--test-file", required=True)
    parser.add_argument("--test-exec-key", default=None)
    parser.add_argument("--arg", default="")
    parser.add_argument("--workspace", default="")
    return parser.parse_args()


def parse_args_other():
    parser = pydantic_argparse.ArgumentParser(
        model=Arguments,
        prog="Example Program",
        description="Example Description",
        version="0.0.1",
    )
    return parser.parse_typed_args()


def main() -> None:
    """Simple Main Function."""
    if platform.system().lower() == "darwin":
        args = parse_args_macos()
    else:
        args = parse_args_other()

    workspace = args.workspace
    tag = args.tag
    test_exec_key = args.test_exec_key
    test_file = args.test_file
    arg = args.arg

    Logs_Directory = CreatFolder(test_file)
    Logs_Directory1 = ""
    if len(workspace) > 0:
        Logs_Directory1 = Logs_Directory
        print("path logs directory", Logs_Directory1)

    if test_file.endswith(".robot"):
        Dossier_Allure = CreatFolderAllure(Logs_Directory1)
        ExcuteRobotTest(Logs_Directory1, Dossier_Allure, test_file, arg, tag)
        AddJenkinsLogToRobot(Logs_Directory)
    elif test_file.endswith(".py"):
        output = ExcutePythonTest(Logs_Directory1, test_file, arg)
        status = ParserLOGPythonTest(output)
        generateJson(Logs_Directory, test_exec_key, tag, status)
    else:
        output = ExcuteNPMTest(test_file)
        print(output)
        status = ParserLOGPythonTest(output)
        generateJson(Logs_Directory, test_exec_key, tag, status)


if __name__ == "__main__":
    main()
