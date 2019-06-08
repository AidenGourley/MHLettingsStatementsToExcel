from calendar import monthrange
from datetime import datetime

import fiscalyear
from termcolor import cprint

import Scraper


class Statement:

    def __init__(self, stringFromPDF=None):
        if stringFromPDF:
            self.incomeTransactions, self.expenditureTransactions = Scraper.updateTransactions(stringFromPDF)
            self.address = Scraper.getAddress(stringFromPDF)
            self.number = Scraper.getStatementNumber(stringFromPDF)
            self.startDate = self.getStatementStartDate(stringFromPDF)
            self.endDate = self.__calcEndDate__()
            self.dateString = self.getStatementDateString(stringFromPDF)
            self.retained = Scraper.getRetained(stringFromPDF)
        else:
            self.incomeTransactions, self.expenditureTransactions = [], []
            self.address = None
            self.number = None
            self.startDate = None
            self.endDate = None
            self.dateString = None
            self.retained = 0

    def __calcEndDate__(self):
        try:
            if self.startDate.month == 12:
                newMonth = 1
                newYear = self.startDate.year + 1
            else:
                newMonth = self.startDate.month + 1
                newYear = self.startDate.year
            if self.startDate.day == 1:
                return fiscalyear.FiscalDate(newYear, newMonth,
                                             monthrange(self.startDate.year, self.startDate.month)[1])
            else:
                return fiscalyear.FiscalDate(newYear, newMonth,
                                             monthrange(newYear, newMonth)[1])  # self.startDate.day - 1)
        except:
            cprint("ERROR: STATEMENT HAS NO DATE ATTRIBUTES! NO. " + self.number, "red")
            return fiscalyear.FiscalDate(2000, 1, 31)

    def getNetIncome(self):
        total = 0
        for transaction in self.incomeTransactions:
            total += transaction.net
        return round(total, 2)

    def getGrossIncome(self):
        total = 0
        for transaction in self.incomeTransactions:
            total += transaction.gross
        return round(total, 2)

    def getNetExpenditure(self):
        total = 0
        for transaction in self.expenditureTransactions:
            total += transaction.net
        return round(total, 2)

    def getGrossExpenditure(self):
        total = 0
        for transaction in self.expenditureTransactions:
            total += transaction.gross
        return round(total, 2)

    def getVATExpenditure(self):
        total = 0
        for transaction in self.expenditureTransactions:
            total += transaction.vat
        return round(total, 2)

    def getGrossProfit(self):
        grossIncome = self.getGrossIncome()
        grossExpenditure = self.getGrossExpenditure()
        grossProfit = grossIncome - grossExpenditure - self.retained
        return round(grossProfit, 2)

    def getNetProfit(self):
        netIncome = self.getNetIncome()
        netExpenditure = self.getNetExpenditure()
        netProfit = netIncome - netExpenditure
        return round(netProfit, 2)

    def getStatementDateString(self, segment):
        string = self.getStatementStartDate(segment).strftime("%d/%m/%Y") + " to " + self.endDate.strftime("%d/%m/%Y")
        return string

    def getStatementStartDate(self, segment): #Returns FiscalDate Obj
        for i in self.incomeTransactions:
            if (("Rent-" in i.name) or ("Rent -" in i.name)) and i.date:
                return i.date
            else:
                try:
                    if ("Rent " + datetime.strftime(i.date, "%d/%m/%Y") in i.name) and i.date:
                        return i.date
                    else:
                        continue
                except:
                    continue
        # Else
        return Scraper.getStatementDate(segment)

    def getGardeningExpenses(self):
        gardenExpenses = []
        for i in self.expenditureTransactions:
            if "Garden" in i.name:
                gardenExpenses.append(i)
        return gardenExpenses

    def getAgentFeeExpenses(self):
        agentFeeExpenses = []
        feeTypes = ["Management Fees", "Inspection", "inspec", "Rent Guarantee", "Check In", "Check Out", "Inventory",
                    "Deposit"]
        for i in self.expenditureTransactions:
            for fee in feeTypes:
                if fee in i.name:
                    agentFeeExpenses.append(i)
                    break
        return agentFeeExpenses

    def getOtherExpenses(self):
        previouslyListedTransactions = self.getAgentFeeExpenses()
        for i in self.getGardeningExpenses():
            previouslyListedTransactions.append(i)
        otherTransactions = []
        for i in self.expenditureTransactions:
            if i not in previouslyListedTransactions:
                otherTransactions.append(i)
        return otherTransactions

    def agentFeesToString(self):
        l = self.getAgentFeeExpenses()
        totalExpense = 0
        string = ""
        for i in l:
            totalExpense += i.gross
            try:
                string += i.date.strftime("%d %b %Y") + " £" + str(round(i.gross, 2)) + " " + i.name + " | "
            except AttributeError:
                string += "£" + str(round(i.gross, 2)) + " " + i.name + " | "
        string += " TOTAL: £" + str(round(totalExpense, 2))
        return string

    def gardeningFeesToString(self):
        l = self.getGardeningExpenses()
        totalExpense = 0
        string = ""
        for i in l:
            totalExpense += i.gross
            try:
                string += i.date.strftime("%d %b %Y") + " £" + str(round(i.gross, 2)) + " " + i.name + " | "
            except AttributeError:
                string += "£" + str(round(i.gross, 2)) + " " + i.name + " | "
        string += " TOTAL: £" + str(round(totalExpense, 2))
        return string

    def otherFeesToString(self):
        l = self.getOtherExpenses()
        totalExpense = 0
        string = ""
        for i in l:
            totalExpense += i.gross
            try:
                string += i.date.strftime("%d %b %Y") + " £" + str(round(i.gross, 2)) + " " + i.name + " | "
            except AttributeError:
                string += "£" + str(round(i.gross, 2)) + " " + i.name + " | "
        string += " TOTAL: £" + str(round(totalExpense, 2))
        return string
