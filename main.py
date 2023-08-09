import pandas as pd
import xlsxwriter


class indexBalancer:
    money_to_invest = 0
    calculation_dataframe = pd.DataFrame
    data = pd.DataFrame

    def get_input(self):
        self.money_to_invest = input('How much would you like to invest?')
        try:
            val = float(self.money_to_invest)
        except:
            print ('Input must be a number. Try Again!')
            self.money_to_invest = input('How much would you like to invest?')
        return

    def initialize_dataframe_columns(self):
        column_values = ['Ticker', 'CurrentPrice', 'MarketCap', 'NumStocksToBuy']
        self.calculation_dataframe = pd.DataFrame(
            columns = column_values
        )
        return

    def retrieve_data(self):
        self.data = pd.read_csv('constituents-financials_csv.csv')
        return

    def parse_data(self):
        my_columns = list(self.calculation_dataframe.columns)
        data_columns = ['Symbol', 'Price', 'Market Cap']
        for i in range(len(my_columns) - 1):
            self.calculation_dataframe[my_columns[i]] = self.data[data_columns[i]]
        return

    def calculate_distribution(self):
        dollars = float(self.money_to_invest)
        total_market_cap = 0
        for caps in self.calculation_dataframe['MarketCap']:
            total_market_cap = total_market_cap + caps
        for i in range(len(self.calculation_dataframe['MarketCap'])):
            proportion = self.calculation_dataframe['MarketCap'][i] / total_market_cap
            num_dollars_to_spend = proportion * dollars
            num_stocks_to_buy = num_dollars_to_spend / self.calculation_dataframe['CurrentPrice'][i]
            self.calculation_dataframe['NumStocksToBuy'][i] = num_stocks_to_buy
        return

    def create_excel(self):
        excelWriter = pd.ExcelWriter('shares_to_buy_s&p500', engine = 'xlsxwriter')
        self.calculation_dataframe.to_excel(excelWriter, 'Shares To Buy', index = False)
        back_color = '#0a0a23'
        font_color = '#fffffe'
        string_format = excelWriter.book.add_format(
            {
                'font_color' : font_color,
                'bg_color' : back_color,
                'border' : 1
            }
        )
        dollar_format = excelWriter.book.add_format(
            {
                'num_format' : '$0.00',
                'font_color': font_color,
                'bg_color': back_color,
                'border': 1
            }
        )
        num_format = excelWriter.book.add_format(
            {
                'num_format' : '0.000',
                'font_color': font_color,
                'bg_color': back_color,
                'border': 1
            }
        )
        format = [['A:A', string_format], ['B:B', dollar_format], ['C:C', dollar_format], ['D:D', num_format]]
        for i in range(len(list(self.calculation_dataframe.columns))):
            excelWriter.sheets['Shares To Buy'].set_column(format[i][0], 18, format[i][1])
        excelWriter.close()
        return


def main():
    balancer = indexBalancer()
    balancer.get_input()
    balancer.initialize_dataframe_columns()
    balancer.retrieve_data()
    balancer.parse_data()
    balancer.calculate_distribution()
    balancer.create_excel()
    return


main()

