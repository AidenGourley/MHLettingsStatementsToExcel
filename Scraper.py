import calendar
import re

import fiscalyear
from termcolor import cprint

import Transaction


def popDates(segment):
    """Finds all the dates within the segment string, removes these dates, and returns the list of dates found, with
    a new segment that does not include the dates."""
    matches = re.findall(r'\d{2}/\d{2}/\d{4}', segment)
    for match in matches:
        segment = segment.replace(match, "")
    return matches, segment


def evaluateStartPosOfNum(segment, startPos):
    while True:
        try:
            int(segment[startPos - 1])
            startPos = startPos - 1
        except ValueError as err:
            if segment[startPos - 1] == "-":
                startPos = startPos - 1
            else:
                break
    return startPos


def validateBalanceCollumn(segment, firstDecimalPos, thirdDecimalPos):
    try:
        int(segment[firstDecimalPos + 2])
        if thirdDecimalPos - firstDecimalPos > 16:
            return False
        return True
    except:
        return False


def removeYearAndMonth(segment, startPos):
    for i in range(1, 13):
        month = calendar.month_name[i]
        if month in segment:
            segment = segment.replace(month, "")
    for i in range(2010, 2030):
        if str(i) in segment:
            beginIndexOfYear = startPos + segment.index(str(i))
            segment = segment[:beginIndexOfYear] + segment[beginIndexOfYear + 4:]
    return segment


def getStartPos(segment, firstDecimalPos):
    netStartPos = firstDecimalPos
    while True:
        try:
            int(segment[netStartPos - 1])
            netStartPos = netStartPos - 1
        except ValueError as err:
            break
    return netStartPos


def getTransactionInfo(segment, startPos, previousEndPos):
    name = segment[previousEndPos + 1: startPos]
    dates, name = popDates(name)
    dateString = ""
    if dates:
        if len(dates) >= 2:
            dateString += dates[0]
            dateString += " to " + dates[1]
        else:
            dateString += dates[0]
    else:
        dateString = None

    if "TOTAL EXPENDITURE" in name:
        name = name.replace("TOTAL EXPENDITURE", "")
        endOfSection = True
    else:
        endOfSection = False
    if dates:
        return name, dates[0], dateString, endOfSection
    else:
        return name, None, dateString, endOfSection


def getTransactionValues(segment, firstDecimalPos):
    # This method will likely produce an inaccurate result for the NET VALUE if another integer precedes it in the segment string.
    secondDecimalPos = segment.find(".", (firstDecimalPos + 1))
    thirdDecimalPos = segment.find(".", (secondDecimalPos + 1))
    # Good
    # Check if the decimal points lie within the expected range, otherwise it is unlikely the segment is
    # part of the balance sheet.

    if not validateBalanceCollumn(segment, firstDecimalPos, thirdDecimalPos):
        return None, None, None, firstDecimalPos, thirdDecimalPos + 3, False

    netStartPos = firstDecimalPos
    netStartPos = evaluateStartPosOfNum(segment, netStartPos)

    non_decimal = re.compile(r'[^\d.-]+')
    VAT = non_decimal.sub("", segment[firstDecimalPos + 3:secondDecimalPos + 3])
    gross = non_decimal.sub("", segment[secondDecimalPos + 3:thirdDecimalPos + 3])
    net = non_decimal.sub("", segment[netStartPos:(firstDecimalPos + 3)])

    cprint("Value Scrape Success: (Net, VAT, Gross)", "green")
    endOfSection = False
    if (segment.find("TOTAL INCOME", thirdDecimalPos + 2) == thirdDecimalPos + 3) or (
            "TOTAL EXPENDITURE" in segment[thirdDecimalPos + 2: thirdDecimalPos + 18]):
        endOfSection = True

    return net, VAT, gross, netStartPos, thirdDecimalPos + 2, endOfSection


def updateTransactions(segment):
    incomeTransactions = []
    expenditureTransactions = []
    endPos = segment.find(".")
    endOfSection = False
    previousEndPos = segment.find("INCOME") + 5
    while not endOfSection:
        firstDecimalPos = segment.find(".", endPos)
        secondDecimalPos = segment.find(".", (firstDecimalPos + 1))
        thirdDecimalPos = segment.find(".", (secondDecimalPos + 1))
        if not validateBalanceCollumn(segment, firstDecimalPos, thirdDecimalPos):
            endPos += 1
            continue
        startPos = getStartPos(segment, firstDecimalPos)
        net, vat, gross, startPos, endPos, endOfSection = getTransactionValues(segment, firstDecimalPos)
        gross = removeYearAndMonth(gross, 0)  ##?
        gross = round(float(gross), 2)
        vat = removeYearAndMonth(vat, 0)  ##?
        vat = round(float(vat), 2)
        net = removeYearAndMonth(net, 0)  ##?
        net = round(float(net), 2)
        expectedNet = gross - vat
        if expectedNet != net:
            cprint("WARNING: Value 'Net' may be wrong: " + str(net), "yellow")
            net = expectedNet
            cprint("AUTO-CORRECTION MADE: Value 'Net' changed to: " + str(net), "yellow")
        name, date, dateString, eos = getTransactionInfo(segment, startPos, previousEndPos)
        previousEndPos = endPos
        if net:
            transaction = Transaction.Transaction(name, net, vat, gross, date, dateString)
            if gross >= 0:
                incomeTransactions.append(transaction)
            else:
                expenditureTransactions.append(transaction)

    endOfSection = False
    newStartPos = endPos
    endPos = segment.find(".", segment.find("EXPENDITURE") + 10)
    previousEndPos = segment.find("EXPENDITURE") + 10
    while not endOfSection:
        firstDecimalPos = segment.find(".", endPos)
        secondDecimalPos = segment.find(".", (firstDecimalPos + 1))
        thirdDecimalPos = segment.find(".", (secondDecimalPos + 1))
        if not validateBalanceCollumn(segment, firstDecimalPos, thirdDecimalPos):
            endPos += 1
            continue
        startPos = getStartPos(segment, firstDecimalPos)
        gross, vat, net, startPos, endPos, endOfSection = getTransactionValues(segment, firstDecimalPos)
        gross = removeYearAndMonth(gross, 0)
        gross = round(float(gross), 2)
        vat = removeYearAndMonth(vat, 0)
        vat = round(float(vat), 2)
        net = removeYearAndMonth(net, 0)
        net = round(float(net), 2)
        expectedNet = gross - vat
        if expectedNet != net:
            cprint("WARNING: Value 'Net' may be wrong: " + str(net), "yellow")
            net = round(expectedNet, 2)
            cprint("AUTO-CORRECTION MADE: Value 'Net' changed to: " + str(net), "yellow")
        name, date, dateString, eos = getTransactionInfo(segment, startPos, previousEndPos)
        if eos:
            endOfSection = True
            net = None
        previousEndPos = endPos
        if net:
            transaction = Transaction.Transaction(name, net, vat, gross, date, dateString)
            if gross >= 0:
                expenditureTransactions.append(transaction)
            else:
                incomeTransactions.append(transaction)
    return incomeTransactions, expenditureTransactions


def getAddress(segment):
    startPos = segment.find("Re: ") + 4
    endPos = segment.find("INCOME")
    segment = segment[startPos:endPos]
    return segment


def getStatementNumber(segment):
    startPosOfNext = segment.find("Statement No")
    startPos = evaluateStartPosOfNum(segment, startPosOfNext)
    return segment[startPos:startPosOfNext]


def getRetained(segment):
    startPos = segment.find("Retained") + 8
    if startPos != 7:
        startPosOfNext = segment.find("BALANCE")
    elif (segment.find("dilapidations") + 12) != 11:
        startPos = segment.find("dilapidations") + 13
        startPosOfNext = segment.find(".", startPos) + 3
    else:
        return 0.00
    retainedStr = segment[startPos:startPosOfNext]
    retained = round(float(retainedStr), 2)
    return retained


def transactionToString(transaction):
    string = transaction.date + " \n" + transaction.name + " \n£" + transaction.gross
    return string


def getStatementDate(segment):
    startPos = segment.find("ADVICE AS AT") + 12
    matches, string = popDates(segment)
    dmy = matches[0].split("/")
    return fiscalyear.FiscalDate(int(dmy[2]), int(dmy[1]), int(dmy[0]))
