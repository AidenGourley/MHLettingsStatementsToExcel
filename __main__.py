import os

import PyPDF2
import fiscalyear  # Change the fiscal year start point in the fiscalyear module
from openpyxl import Workbook
from openpyxl import load_workbook
from termcolor import cprint

import Statement
import Transaction


def getFiscalMonth(statement):
    return ((statement.startDate.month - 4) % 12) + 1


def splitDate(startDate, endDate):
    endStatementStartDate = startDate
    endStatementEndDate = fiscalyear.FiscalDate(year=endDate.year, month=4, day=5)
    startStatementStartDate = fiscalyear.FiscalDate(year=endDate.year, month=4, day=6)
    startStatementEndDate = endDate
    return endStatementStartDate, endStatementEndDate, startStatementStartDate, startStatementEndDate


def splitStatement(statement, Houses):
    # Split the statements and assign new dates first.
    endStatement, startStatement = Statement.Statement(), Statement.Statement()
    endStatement.address, startStatement.address = statement.address, statement.address
    endStatement.startDate, endStatement.endDate, startStatement.startDate, startStatement.endDate = splitDate(
        statement.startDate, statement.endDate)
    differenceBetweenDates = abs((statement.endDate - statement.startDate).days)
    daysUntilNewFiscalYear = abs((endStatement.endDate - endStatement.startDate).days)
    endStatement.retained = round((statement.retained / differenceBetweenDates) * daysUntilNewFiscalYear, 2)
    startStatement.retained = round(
        (statement.retained / differenceBetweenDates) * (differenceBetweenDates - daysUntilNewFiscalYear), 2)
    startStatement.dateString = startStatement.startDate.strftime(
        "%d/%m/%Y") + " to " + startStatement.endDate.strftime("%d/%m/%Y")
    endStatement.dateString = endStatement.startDate.strftime("%d/%m/%Y") + " to " + endStatement.endDate.strftime(
        "%d/%m/%Y")

    # Remove retained transaction from the endStatement, and keeps it in the startStatement
    for transaction in statement.incomeTransactions:
        if "refund" in transaction.name.lower():
            endStatement.incomeTransactions.append(transaction)
        else:
            # Split the transaction and assign it a weight. Add each new transaction back to the new statements.
            endNet = round((transaction.net / differenceBetweenDates) * daysUntilNewFiscalYear, 2)
            endGross = round((transaction.gross / differenceBetweenDates) * daysUntilNewFiscalYear, 2)
            endVAT = round((transaction.VAT / differenceBetweenDates) * daysUntilNewFiscalYear, 2)
            startNet = round(
                (transaction.net / differenceBetweenDates) * (differenceBetweenDates - daysUntilNewFiscalYear), 2)
            startGross = round((transaction.gross / differenceBetweenDates) * (
                    differenceBetweenDates - daysUntilNewFiscalYear), 2)
            startVAT = round(
                (transaction.VAT / differenceBetweenDates) * (differenceBetweenDates - daysUntilNewFiscalYear), 2)
            splitEndTransaction = Transaction.Transaction(transaction.name, endNet, endVAT, endGross, transaction.date)
            splitStartTransaction = Transaction.Transaction(transaction.name, startNet, startVAT, startGross,
                                                            transaction.date)
            endStatement.incomeTransactions.append(splitEndTransaction)
            startStatement.incomeTransactions.append(splitStartTransaction)

    for transaction in statement.expenditureTransactions:
        endStatement.expenditureTransactions.append(transaction)

    endStatement.number, startStatement.number = statement.number

    if endStatement.endDate.fiscal_year in Houses[house]:  # If the Year record exists then...
        Houses[house][endStatement.endDate.fiscal_year].append(endStatement)
    else:  # Otherwise create a record for a new year...
        Houses[house][endStatement.endDate.fiscal_year] = [endStatement]

    if startStatement.endDate.year in Houses[house]:
        Houses[house][startStatement.endDate.fiscal_year].append(startStatement)
    else:
        Houses[house][startStatement.endDate.fiscal_year] = [startStatement]

    return Houses


Houses = {"RG45 6RY": {}, "RG45 6JS": {}}
statements = []
for filename in os.listdir("./testFiles/"):
    if filename.endswith(".pdf"):
        filePath = "./testFiles/" + filename
        pdfFileObj = open(filePath, 'rb')
        pdfReader = PyPDF2.PdfFileReader(pdfFileObj)
        stringFromPDF = ""
        for i in range(0, pdfReader.numPages):
            pageObj = pdfReader.getPage(i)
            stringFromPDF += pageObj.extractText()
        statement = Statement.Statement(stringFromPDF)
        for house in Houses:  # for each house
            if house in statement.address:  # if the statement is for this house then...
                if statement.startDate.fiscal_year != statement.endDate.fiscal_year:  # if the month is april then.. Split the Statement
                    Houses = splitStatement(statement, Houses)
                    break
                else:
                    if statement.endDate.fiscal_year in Houses[house]:  # If the Year record exists then...
                        Houses[house][statement.endDate.fiscal_year].append(statement)
                    else:  # Otherwise create a record for a new year...
                        Houses[house][statement.endDate.fiscal_year] = [statement]

for house in Houses:
    for year in Houses[house]:
        AnnualGrossProfit = 0
        AnnualRetained = 0
        AnnualAgentFees = 0
        AnnualGardeningFees = 0
        AnnualOtherExpenses = 0
        AnnualExpenditure = 0
        AnnualGrossIncome = 0
        Houses[house][year].sort(key=lambda s: s.startDate)
        for s in Houses[house][year]:
            ##GETS VALUES OF TOTAL COLUMNS ##
            for transaction in s.getAgentFeeExpenses():
                AnnualAgentFees += transaction.gross
            for transaction in s.getGardeningExpenses():
                AnnualGardeningFees += transaction.gross
            for transaction in s.getOtherExpenses():
                AnnualOtherExpenses += transaction.gross
            AnnualGrossIncome += s.getGrossIncome()
            AnnualGrossProfit += s.getGrossProfit()
            AnnualExpenditure += s.getGrossExpenditure()
            AnnualRetained += s.retained
            ##################################
            wbName = house + ".xlsx"
            try:
                wb = load_workbook(wbName)
                try:
                    ws = wb.get_sheet_by_name(str(s.endDate.fiscal_year))

                except:
                    for i in wb.get_sheet_names():
                        if "Sheet" in i:
                            ws = wb.get_sheet_by_name(i)
                            ws.title = str(s.endDate.fiscal_year)
                            ws.cell(row=1, column=1).value = "FiscalYear: " + str(s.startDate.fiscal_year)
                            ws.cell(row=2, column=1).value = "Month"
                            ws.cell(row=2, column=2).value = "Gross Income"
                            ws.cell(row=2, column=3).value = "Agent Fees"
                            ws.cell(row=2, column=4).value = "Gardening Expenditure"
                            ws.cell(row=2, column=5).value = "Other Expenditure"
                            ws.cell(row=2, column=6).value = "Total Expenditure"
                            ws.cell(row=2, column=7).value = "Retained"
                            ws.cell(row=2, column=8).value = "Gross Profit"
                            break

            except:
                wb = Workbook(wbName)
                ws = wb.get_sheet_by_name(str(s.endDate.fiscal_year))
                cprint("WRITE ERROR: NO WORKBOOK TO USE FOR HOUSE", "red")

            row = getFiscalMonth(s) + 2
            ws.cell(row=row, column=1).value = s.dateString
            ws.cell(row=row, column=2).value = s.getGrossIncome()
            ws.cell(row=row, column=3).value = s.agentFeesToString()
            ws.cell(row=row, column=4).value = s.gardeningFeesToString()
            ws.cell(row=row, column=5).value = s.otherFeesToString()
            ws.cell(row=row, column=6).value = s.getGrossExpenditure()
            ws.cell(row=row, column=7).value = s.retained
            ws.cell(row=row, column=8).value = s.getGrossProfit()
            # wb[Houses[house]].append()
            ws.cell(row=15, column=1).value = "TOTAL: "
            ws.cell(row=15, column=2).value = AnnualGrossIncome
            ws.cell(row=15, column=3).value = AnnualAgentFees
            ws.cell(row=15, column=4).value = AnnualGardeningFees
            ws.cell(row=15, column=5).value = AnnualOtherExpenses
            ws.cell(row=15, column=6).value = AnnualExpenditure
            ws.cell(row=15, column=7).value = AnnualRetained
            ws.cell(row=15, column=8).value = AnnualGrossProfit
            print(s.startDate, " - ", s.endDate)

            wb.save(wbName)

cprint("Spreadsheet Build Success!", "green")
input()
