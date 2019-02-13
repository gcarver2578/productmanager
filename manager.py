# Gabriel Carver, Holden Stringfield, Eli Horton
# This program is for analyzing data on products and potential sales
import graphics as g
import pandas as pd


def main():
    prediction()
  
  
def prediction():
    p_win = g.GraphWin('Prediction Graph', 100, 100)
    p_win.close()
    file = 'Products.xlsx'
    data = pd.read_excel(file)
    # data.set_index('2017', inplace=True)
    # print(list(data.loc['Paper', :].values))
    # new_val = [list(data.loc[x, :].values) for x in range(10)]
    # print(new_val)
    print(data)
    return


main()
