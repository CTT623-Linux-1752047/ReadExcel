import time

from neo4j import GraphDatabase
import const
import common


class databaseHelpers:

    # constructor of the class for create connect with neo4j DB
    def __init__(self, uri: object, user: object, password: object):
        try :
            self.driver = GraphDatabase.driver(uri, auth=(user, password), max_connection_lifetime=1000)

        except Exception as e:
            common.writeLog(e)


    # close connect with neo4j DB
    def close(self):
        # if self.__driver is not None:
            self.driver.close()

    # function check connect to database is exist, if no return 0
    def isConnecting(self, databaseName):
        try:
            self.driver.verify_connectivity(database=databaseName)
            return True;
        except:
            return False;

    def getAllNode(self, db=None):
        assert self.driver is not None, "Driver not initialized!"
        nodes = []
        session = None
        try:
            session = self.driver.session(database=db) if db is not None else self.driver.session()
            response = list(session.run("MATCH (n) RETURN n"))
            # convert data
            for record in response:
                if "type" in record['n'] :
                    nodes.append({"code": record['n']['code'], "value": record['n']['value'], "type": record['n']['type']})
                else :
                    nodes.append({"code": record['n']['code'], "nom": record['n']['nom']})
        except Exception as e:
            common.writeLog(e)
        finally:
            if session is not None:
                session.close()

        return nodes

    #function is create node Variable - if create fail return 0 <> return ...
    def createVariable(self, lstVariable, db=None ):
        result = 0;

        try:
            if self.isConnecting(db) :
                session = self.driver.session(database=db) if db is not None else self.driver.session()
                for item in lstVariable:
                    if item.label is None:
                        query = f"MERGE (v:{const.NAME_VARIABLE} {const.CURLY_BRACKETS_OPEN} variableCd: \"{item.code}\" , variableDescription: \"{item.value}\"{const.CURLY_BRACKETS_CLOSE} ) RETURN v"
                    else:
                        query = f"MERGE (v:{const.NAME_VARIABLE} {const.CURLY_BRACKETS_OPEN} variableCd: \"{item.code}\" , variableDescription: \"{item.value}\"{const.CURLY_BRACKETS_CLOSE} ) RETURN v"

                    result = session.run(query)
                    print(query)
        except Exception as e:
            common.writeLog("databaseHelpers > createVariable : "  + str(e))
        return result;

    # function is create node Variable Value - if create fail return 0 <> return ...
    def createVariableValue(self, lstVariableValues, db=None ):
        result = 0;

        try:
            if self.isConnecting(db) :
                session = self.driver.session(database=db) if db is not None else self.driver.session()
                for item in lstVariableValues:
                    if item.label is None:
                        query = f"MERGE (vv:{const.NAME_VARIABLE_VALUE} {const.CURLY_BRACKETS_OPEN} variableValueCd: \"{item.code}\" , variableValue: \"{item.value}\", variableType: \"{item.type}\" {const.CURLY_BRACKETS_CLOSE} ) RETURN vv"
                    else:
                        query = f"MERGE (vv:{const.NAME_VARIABLE_VALUE} {const.CURLY_BRACKETS_OPEN} variableValueCd: \"{item.code}\" , variableValue: \"{item.value}\", variableType: \"{item.type}\" {const.CURLY_BRACKETS_CLOSE} ) RETURN vv"

                    result = session.run(query)
                    print(query)
        except Exception as e:
            common.writeLog("databaseHelpers > createVariableValue : "  + str(e))
        return result;

    # function is create node Technique - if create fail return 0 <> return ...
    def createTechnique(self, lstTechniques, db=None ):
        result = 0;

        try:
            if self.isConnecting(db) :
                session = self.driver.session(database=db) if db is not None else self.driver.session()
                for item in lstTechniques:
                    if item.label is None:
                        query = f"MERGE (t:{const.NAME_TECHNIQUE} {const.CURLY_BRACKETS_OPEN} techniqueCd: \"{item.code}\" , techniqueName: \"{item.value}\" {const.CURLY_BRACKETS_CLOSE} ) RETURN t"
                    else:
                        query = f"MERGE (t:{const.NAME_TECHNIQUE} {const.CURLY_BRACKETS_OPEN} techniqueCd: \"{item.code}\" , techniqueName: \"{item.value}\" {const.CURLY_BRACKETS_CLOSE} ) RETURN t"

                    result = session.run(query)
                    print(query)
        except Exception as e:
            common.writeLog("databaseHelpers > createTechnique : "  + str(e))
        return result;

    # function is create node Student - if create fail return 0 <> return ...
    def createStduent(self, lstStudents, db=None):
        result = 0;

        try:
            if self.isConnecting(db):
                session = self.driver.session(database=db) if db is not None else self.driver.session()
                for item in lstStudents:
                    if item.label is None:
                        query = f"MERGE (st:{const.NAME_STUDENT} {const.CURLY_BRACKETS_OPEN} studentCd: \"{item.code}\" , fullNameStudent: \"{item.value}\" {const.CURLY_BRACKETS_CLOSE} ) RETURN st"
                    else:
                        query = f"MERGE (st:{const.NAME_STUDENT} {const.CURLY_BRACKETS_OPEN} studentCd: \"{item.code}\" , fullNameStudent: \"{item.value}\" {const.CURLY_BRACKETS_CLOSE} ) RETURN st"

                    result = session.run(query)
                    print(query)
        except Exception as e:
            common.writeLog("databaseHelpers > createStduent : " + str(e))
        return result;

    # function is create node Group - if create fail return 0 <> return ...
    def createGroup(self, lstGroups, db=None):
        result = 0;

        try:
            if self.isConnecting(db):
                session = self.driver.session(database=db) if db is not None else self.driver.session()
                for item in lstGroups:
                    if item.label is None:
                        query = f"MERGE (g:{const.NAME_GROUP} {const.CURLY_BRACKETS_OPEN} groupCd: \"{item.code}\" , groupName: \"{item.value}\" {const.CURLY_BRACKETS_CLOSE} ) RETURN g"
                    else:
                        query = f"MERGE (g:{const.NAME_GROUP} {const.CURLY_BRACKETS_OPEN} groupCd: \"{item.code}\" , groupName: \"{item.value}\" {const.CURLY_BRACKETS_CLOSE} ) RETURN g"

                    result = session.run(query)
                    print(query)
        except Exception as e:
            common.writeLog("databaseHelpers > createGroup : " + str(e))
        return result;

    # function is create node Level - if create fail return 0 <> return ...
    def createLevel(self, lstLevels, db=None):
        result = 0;

        try:
            if self.isConnecting(db):
                session = self.driver.session(database=db) if db is not None else self.driver.session()
                for item in lstLevels:
                    if item.label is None:
                        query = f"MERGE (l:{const.NAME_LEVEL} {const.CURLY_BRACKETS_OPEN} levelCd: \"{item.code}\" , levelName: \"{item.value}\" {const.CURLY_BRACKETS_CLOSE} ) RETURN l"
                    else:
                        query = f"MERGE (l:{const.NAME_LEVEL} {const.CURLY_BRACKETS_OPEN} levelCd: \"{item.code}\" , levelName: \"{item.value}\" {const.CURLY_BRACKETS_CLOSE} ) RETURN l"

                    result = session.run(query)
                    print(query)

        except Exception as e:
            common.writeLog("databaseHelpers > createLevel : " + str(e))
        return result;

    # function is create node LocalInsInfo - if create fail return 0 <> return ...
    def createLocalInsInfo(self, lstLocalInsInfo, db=None):
        result = 0;

        try:
            if self.isConnecting(db):
                session = self.driver.session(database=db) if db is not None else self.driver.session()
                for item in lstLocalInsInfo:
                    if item.label is None:
                        query = f"MERGE (li:{const.NAME_LOCALINSINFO} {const.CURLY_BRACKETS_OPEN} localInsInfoCd: \"{item.code}\"  {const.CURLY_BRACKETS_CLOSE} ) RETURN li"
                    else:
                        query = f"MERGE (li:{const.NAME_LOCALINSINFO} {const.CURLY_BRACKETS_OPEN} localInsInfoCd: \"{item.code}\"  {const.CURLY_BRACKETS_CLOSE} ) RETURN li"

                    result = session.run(query)
                    print(query)

        except Exception as e:
            common.writeLog("databaseHelpers > createLocalInsInfo : " + str(e))
        return result;

    def runQuery(self, query, db):
        result = 0;
        try:
            session = self.driver.session(database=db) if db is not None else self.driver.session()
            print(self.driver.session())
            print(session)
            result = session.run(query)
        except Exception as e:
            common.writeLog("databaseHelpers > runQuery : " + str(e))
        return result;

    # function is create node Task Type - if create fail return 0 <> return ...
    def createTaskType(self, lstTaskTypes , db=None ):
        result = 0;

        try:
            if self.isConnecting(db) :
                session = self.driver.session(database=db) if db is not None else self.driver.session()
                for item in lstTaskTypes:
                    if item.label is None:
                        query = f"MERGE (tt:{const.NAME_TASK_TYPE} {const.CURLY_BRACKETS_OPEN} taskTypeCd: \"{item.code}\" , taskTypeName: \"{item.name}\" {const.CURLY_BRACKETS_CLOSE} ) "
                    else:
                        query = f"MERGE (tt:{const.NAME_TASK_TYPE} {const.CURLY_BRACKETS_OPEN} taskTypeCd: \"{item.code}\" , taskTypeName: \"{item.name}\" {const.CURLY_BRACKETS_CLOSE} ) "

                    result = session.run(query)
                    print(query)
        except Exception as e:
            common.writeLog("databaseHelpers > createTaskType : "  + str(e))
        return result;

    # function is create relationship between 2 node
    def createRelationship(self, lstRelationships, db=None):
        result = 0;

        try:
            session = self.driver.session(database=db) if db is not None else self.driver.session()
            if self.isConnecting(db) :
                for item in lstRelationships:
                    if(item.typeL == const.NAME_TECHNIQUE) :
                        query1 = f" MATCH (left {const.CURLY_BRACKETS_OPEN} techniqueCd:\"{item.left}\" {const.CURLY_BRACKETS_CLOSE})"
                    elif (item.typeL == const.NAME_TASK_TYPE) :
                        query1 = f" MATCH (left {const.CURLY_BRACKETS_OPEN} taskTypeCd:\"{item.left}\" {const.CURLY_BRACKETS_CLOSE})"
                    elif (item.typeL == const.NAME_VARIABLE) :
                        query1 = f" MATCH (left {const.CURLY_BRACKETS_OPEN} variableCd:\"{item.left}\" {const.CURLY_BRACKETS_CLOSE})"
                    elif (item.typeL == const.NAME_GROUP) :
                        query1 = f" MATCH (left {const.CURLY_BRACKETS_OPEN} groupCd:\"{item.left}\" {const.CURLY_BRACKETS_CLOSE})"
                    elif (item.typeL == const.NAME_LEVEL):
                        query1 = f" MATCH (left {const.CURLY_BRACKETS_OPEN} levelCd:\"{item.left}\" {const.CURLY_BRACKETS_CLOSE})"
                    elif (item.typeL == const.NAME_STUDENT):
                        query1 = f" MATCH (left {const.CURLY_BRACKETS_OPEN} studentCd:\"{item.left}\" {const.CURLY_BRACKETS_CLOSE})"
                    elif (item.typeL == const.NAME_LOCALINSINFO):
                        query1 = f" MATCH (left {const.CURLY_BRACKETS_OPEN} localInsInfoCd:\"{item.left}\" {const.CURLY_BRACKETS_CLOSE})"
                    else :
                        query1 = f" MATCH (left {const.CURLY_BRACKETS_OPEN} variableValueCd:\"{item.left}\" {const.CURLY_BRACKETS_CLOSE})"

                    if (item.typeR == const.NAME_TECHNIQUE):
                        query2 = f" MATCH (right {const.CURLY_BRACKETS_OPEN} techniqueCd:\"{item.right}\" {const.CURLY_BRACKETS_CLOSE})"
                    elif (item.typeR == const.NAME_TASK_TYPE):
                        query2 = f" MATCH (right {const.CURLY_BRACKETS_OPEN} taskTypeCd:\"{item.right}\" {const.CURLY_BRACKETS_CLOSE})"
                    elif (item.typeR == const.NAME_VARIABLE):
                        query2 = f" MATCH (right {const.CURLY_BRACKETS_OPEN} variableCd:\"{item.right}\" {const.CURLY_BRACKETS_CLOSE})"
                    elif (item.typeR == const.NAME_GROUP):
                        query2 = f" MATCH (right {const.CURLY_BRACKETS_OPEN} groupCd:\"{item.right}\" {const.CURLY_BRACKETS_CLOSE})"
                    elif (item.typeR == const.NAME_LEVEL):
                        query2 = f" MATCH (right {const.CURLY_BRACKETS_OPEN} levelCd:\"{item.right}\" {const.CURLY_BRACKETS_CLOSE})"
                    elif (item.typeR == const.NAME_STUDENT):
                        query2 = f" MATCH (right {const.CURLY_BRACKETS_OPEN} studentCd:\"{item.right}\" {const.CURLY_BRACKETS_CLOSE})"
                    elif (item.typeR == const.NAME_LOCALINSINFO):
                        query2 = f" MATCH (right {const.CURLY_BRACKETS_OPEN} localInsInfoCd:\"{item.right}\" {const.CURLY_BRACKETS_CLOSE})"
                    else:
                        query2 = f" MATCH (right {const.CURLY_BRACKETS_OPEN} variableValueCd:\"{item.right}\" {const.CURLY_BRACKETS_CLOSE})"

                    query3 = f" MERGE (left)-[r:{item.nameRelationship}]->(right);"

                    result = session.run(query1 + query2 + query3)
                    print(query1 + query2 + query3)

        except Exception as e:
            common.writeLog(e)
        return result;

    def getResultCreate(self, result):
        if len([record.data() for record in result]) > 0:
            print([record.data() for record in result])
            return [record.data() for record in result]
        else:
            return False




