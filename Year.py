import fiscalyear
fiscalyear.START_DAY = 6
fiscalyear.START_MONTH = 4


class Year:
    def __init__(self, year):
        self.__year = year
        self.__statements = []
        self.grossProfit = 0
        self.retained = 0
        self.agentFees = 0
        self.gardeningFees = 0
        self.otherExpenses = 0
        self.expenditure = 0
        self.grossIncome = 0

    def getStatements(self):
        return self.__statements

    def addStatement(self, statement):
        self.__statements.append(statement)
        self.__statements.sort(key=lambda s: s.startDate)

    def removeStatement(self, statementStartDate):
        for statement in self.__statements:
            if statement.startDate == statementStartDate:
                self.__statements.remove(statement)
                print("Statement Removed")
                return None
        print("Unable to find and remove the statement")

    def getYear(self):
        return self.__year

    def totalGrossProfit(self):
        grossProfit = 0
        for statement in self.getStatements():
            grossProfit += statement.getGrossProfit()
        return grossProfit

    def totalGrossIncome(self):
        grossIncome = 0
        for statement in self.getStatements():
            grossIncome += statement.getGrossIncome()
        return grossIncome

    def totalRetained(self):
        retained = 0
        for statement in self.getStatements():
            retained += statement.retained
        return retained

    def totalAgentFees(self):
        agentFees = 0
        for statement in self.getStatements():
            for transaction in statement.getGardeningExpenses():
                agentFees += transaction.gross
        return agentFees

    def totalGardeningFees(self):
        gardeningFees = 0
        for statement in self.getStatements():
            for transaction in statement.getGardeningExpenses():
                gardeningFees += transaction.gross
        return gardeningFees

    def totalOtherExpenditure(self):
        otherExpenses = 0
        for statement in self.getStatements():
            for transaction in statement.getOtherExpenses():
                otherExpenses += transaction.gross
        return otherExpenses

    def totalGrossExpenditure(self):
        grossExpenditure = 0
        for statement in self.getStatements():
            grossExpenditure += statement.getGrossExpenditure()
        return grossExpenditure

