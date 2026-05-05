from Lib_RuneurJenkins import *
import sys
import pydantic
import pydantic_argparse


class Arguments(pydantic.BaseModel):
    """Simple Command-Line Arguments."""
    # Required Args
    tag: str = pydantic.Field(description="a required string",default=None)
    test_file: str = pydantic.Field(description="a required integer")
    test_exec_key: str = pydantic.Field(description="a required integer",default=None)
    arg: str = pydantic.Field(description="a required integer",default="")
    workspace: str = pydantic.Field(description="a required integer",default="")
    
def main() -> None:
    """Simple Main Function."""
    # Create Parser and Parse Args
    parser = pydantic_argparse.ArgumentParser(
        model=Arguments,
        prog="Example Program",
        description="Example Description",
        version="0.0.1", 
    )
    args = parser.parse_typed_args()
    workspace =args.workspace
    tag = args.tag
    test_exec_key= args.test_exec_key 
    test_file = args.test_file
    arg = args.arg
    
    # Print Args
    
    Logs_Directory = CreatFolder(test_file)
    Logs_Directory1=""
    if len(workspace)>0:
       Logs_Directory1 = Logs_Directory
       print("path logs directory",Logs_Directory1)

    if test_file.endswith(".robot"):
    
        Dossier_Allure = CreatFolderAllure(Logs_Directory1)
        
        ExcuteRobotTest(Logs_Directory1, Dossier_Allure, test_file,arg,tag)
        AddJenkinsLogToRobot(Logs_Directory)
    elif test_file.endswith(".py"):   
        output=ExcutePythonTest(Logs_Directory1, test_file,arg)
        status=ParserLOGPythonTest(output)
        generateJson(Logs_Directory,test_exec_key,tag,status) 
    else :   
        output=ExcuteNPMTest(test_file)
        print(output)
        status=ParserLOGPythonTest(output)
        generateJson(Logs_Directory,test_exec_key,tag,status) 
         
    



if __name__ == "__main__":
    main()
