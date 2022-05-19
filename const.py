from datetime import date
from datetime import datetime

CURLY_BRACKETS_OPEN = "{"
CURLY_BRACKETS_CLOSE = "}"
DATE_CD_DMYHMS = date.today().strftime("%d%m%Y%H%M%S")
DATE_DMYHMS = datetime.now().strftime("%d/%m/%Y %H:%M:%S")

SHEET_NAME_VARIABLE = "Variable"
SHEET_NAME_VARIABLE_VALUE = "Variable Value"
SHEET_NAME_TASK_TYPE ="Task Type"
SHEET_NAME_TECHNIQUE = "Technique"
SHEET_NAME_LEVEL = "Level"
SHEET_NAME_STUDENT = "Student"
SHEET_NAME_GROUP = "Group"
SHEET_NAME_LOCALINSINFO = "LocalInsInfo"

NAME_LEVEL = "Level"
NAME_STUDENT = "Student"
NAME_GROUP = "Group"
NAME_LOCALINSINFO = "LocalInsInfo"
NAME_VARIABLE = "Variable"
NAME_VARIABLE_VALUE = "VariableValue"
NAME_TASK_TYPE ="TaskType"
NAME_TECHNIQUE = "Technique"

RELATIONSHIP_VARIABLEVALUE_VARIABLEVALUE = "hasSubValue"
RELATIONSHIP_VARIABLEVALUE_VARIABLE = "hasVariableType"
RELATIONSHIP_TASKTYPE_TASKTYPE = "hasSubTaskType"
RELATIONSHIP_TASKTYPE_VARIABLEVALUE = "hasVariableValue"
RELATIONSHIP_TASKTYPE_TECHNIQUE = "hasPragmaticScope"
RELATIONSHIP_TT_GLOBAL = "hasGlobalConcurrency"
RELATIONSHIP_TT_INC = "hasPartialConcurrency"
RELATIONSHIP_TT_EXT = "hasConcurrencyByExtension"
RELATIONSHIP_STUDENT_GROUP = "arrangedIn"
RELATIONSHIP_LOCALINSINFO_STUDENT = "relatedToStudent"
RELATIONSHIP_LOCALINSINFO_LEVEL = "relatedToLevel"
RELATIONSHIP_LOCALINSINFO_TECHNIQUE = "relatedToTech"

SUB_TASK = "SubTask"

MSG_001 = "sheet {0} : Column A : Row {1} has empty value. "
MSG_002 = "Column of variable'Task Type is not macth with row of sheet Variable. "
MSG_003 = "Sub task type {0} isn't exist with task type. "
MSG_004 = "Code variable {0} of sheet Variable Value isn't exist in code of sheet Variable. "
MSG_005 = "Code task type {0} of sheet Technique isn't exist in code of sheet Task Type. "
MSG_006 = "Code technique concurency global {0} isn't exist in code of sheet Technique. "
MSG_007 = "Code technique concurency partial {0} isn't exist in code of sheet Technique. "
MSG_008 = "Code technique concurency extension {0} isn't exist in code of sheet Technique. "
MSG_009 = "Code Group {0} of sheet Student isn't exist in code of sheet Group. "
MSG_010 = "Cell {0} has code of Student in sheet LocalInsInfo isn't exist code in of sheet Student. "
MSG_011 = "Cell {0} has code of Level in sheet LocalInsInfo isn't exist code in of sheet Level. "
MSG_012 = "Cell {0} has code of Technique in sheet LocalInsInfo isn't exist code in of sheet Technique."