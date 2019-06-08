import Year


class House:

    def __init__(self, postcode):
        self.__postcode = postcode
        self. __years = []


    def addStatement(self, statement):
        fiscalYear = statement.endDate.fiscal_year
        for year in self.__years:
            if year.getYear() == fiscalYear:
                year.addStatement(statement)
                print("Added successfully")
                return None
        newYear = Year.Year(fiscalYear)
        newYear.addStatement(statement)
        self.__years.append(newYear)
        self.__years.sort(key=lambda y: y.getYear())
        print("Added successfully")

    def removeStatement(self, statementStartDate):
        for year in self.__years:
            if year.__year == statementStartDate.fiscal_year:
                year.removeStatement(statementStartDate)

    def getStatementsByYear(self, fiscalYear):
        for year in self.__years:
            if year.__year == fiscalYear:
                return year.getStatements()
        print("Could not get any statements from that year")

    def getPostcode(self):
        return self.__postcode

    def getYears(self):
        listOfYears = []
        for year in self.__years:
            listOfYears.append(year)
        return listOfYears