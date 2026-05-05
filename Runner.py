import platform

system_name = platform.system()

if system_name == 'Windows':
    ROBOT_WORK_DIR = r'C:\Cubestudio_AutoValid\source\Main_scenarios\RFW'
    PYTHON_WORK_DIR = r'C:\Cubestudio_AutoValid\source\Main_scenarios\Python'
    NPM_WORK_DIR = r'C:\Cubestudio_AutoValid\source\utilities\CoopAPI\CoopApi'
    base_temp = r'C:\Temp'

elif system_name == 'Darwin':
    ROBOT_WORK_DIR = '/Users/tuncubevalidadm/Documents/Cubestudio_AutoValid/source/Main_scenarios/RFW'
    PYTHON_WORK_DIR = '/Users/tuncubevalidadm/Documents/Cubestudio_AutoValid/source/Main_scenarios/Python'
    NPM_WORK_DIR = '/Users/tuncubevalidadm/Documents/Cubestudio_AutoValid/source/utilities/CoopAPI/CoopApi'
    base_temp = '/tmp'

else:
    ROBOT_WORK_DIR = '/local/home/tuncube2/Cubestudio_AutoValid/source/Main_scenarios/RFW'
    PYTHON_WORK_DIR = '/local/home/tuncube2/Cubestudio_AutoValid/source/Main_scenarios/Python'
    NPM_WORK_DIR = '/local/home/tuncube2/Cubestudio_AutoValid/source/utilities/CoopAPI/CoopApi'
    base_temp = '/tmp'
