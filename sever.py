import const
from ExcelData import ExcelData
from databaseHelpers import databaseHelpers
import sys

if __name__ == '__main__':
    excel = ExcelData('C:/Users/DavidLu/Desktop/DataSample.xlsx')
    greeter = databaseHelpers("bolt://localhost:7687", "neo4j", "123")

    if( excel.validateFileExcel()) :
        data = excel.readExcel()


        greeter.createVariable(data['Node'][const.NAME_VARIABLE],'neo4j')
        greeter.createVariableValue(data['Node'][const.NAME_VARIABLE_VALUE], 'neo4j')
        greeter.createTaskType(data['Node'][const.NAME_TASK_TYPE], 'neo4j')
        greeter.createTechnique(data['Node'][const.NAME_TECHNIQUE], 'neo4j')
        greeter.createStduent(data['Node'][const.NAME_STUDENT], 'neo4j')
        greeter.createGroup(data['Node'][const.NAME_GROUP], 'neo4j')
        greeter.createLevel(data['Node'][const.NAME_LEVEL], 'neo4j')
        greeter.createLocalInsInfo(data['Node'][const.NAME_LOCALINSINFO], 'neo4j')
        greeter.createRelationship(data['Relationship'], 'neo4j')

    greeter.close()






