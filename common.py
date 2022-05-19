import logging
import const


def writeLog(message):
    logging.basicConfig(filename=f"{const.DATE_CD_DMYHMS}_log.txt", encoding='utf-8')
    logging.error(f" {const.DATE_DMYHMS} - :  {message}")


def writeNotification(codeMessage, lstParam=[]):
    logging.basicConfig(filename=f"{const.DATE_CD_DMYHMS}_log.txt", encoding='utf-8')
    message = codeMessage
    if lstParam :
        cnt = 0
        for param in lstParam :
            message = message.replace(const.CURLY_BRACKETS_OPEN + str(cnt) + const.CURLY_BRACKETS_CLOSE, param)
            cnt =+1
    logging.warning(f" {const.DATE_DMYHMS} - {message}")

def generateCyperQueryToTxt(str):
    f = open("myfile.txt", "r")
    str = f.read() + "\n" + str
    f.close()

    f = open("myfile.txt", "w")
    f.write(str)
    f.close()
