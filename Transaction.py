import fiscalyear


class Transaction:

    def __init__(self, name="", net=0, VAT=0, gross=0, date=None, dateString=""):
        self.name = name
        self.net = abs(round(net, 2))
        self.VAT = abs(round(VAT, 2))
        self.gross = abs(round(gross, 2))
        self.dateString = dateString
        try:
            self.date = fiscalyear.FiscalDate(int(date[6:10]), int(date[3:5]), int(date[:2]))
        except:
            self.date = date
