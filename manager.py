# Gabriel Carver, Holden Stringfield, and Eli Horton
# This program is for analyzing data on products and potential sales
import graphics as g
import pandas as pd


def main():
    prediction()
  
  
def prediction():
    p_win = g.GraphWin('Home', 500, 400)
    p_win.setBackground('white')
    file = 'Products.xlsx'
    data = pd.read_excel(file)
    inp = g.Entry(g.Point(50, 385), 10)
    inp.draw(p_win)
    t = g.Text(g.Point(250, 150), "Type in the product you want to view and then press 'Graph'.")
    t.draw(p_win)
    _close_ = g.Rectangle(g.Point(450, 0), g.Point(500, 50))
    close_text = g.Text(g.Point(475, 25), 'Exit')
    _close_.setFill('red')
    _close_.draw(p_win)
    _graph_ = g.Rectangle(g.Point(450, 50), g.Point(500, 100))
    _graph_.setFill("lightgreen")
    _graph_.draw(p_win)
    graph_text = g.Text(g.Point(475, 75), 'Graph')
    for item in [close_text, graph_text]:
        item.setFill('white')
        item.setStyle('bold')
        item.draw(p_win)
    m = p_win.getMouse()
    while not inside([m], _close_.getP1(), _graph_.getP2()):
        m = p_win.getMouse()
    if inside([m], _close_.getP1(), _close_.getP2()):
        p_win.close()
        return
    product = inp.getText().title().strip()
    col = list(data.loc[:, 2016].values)
    try:
        ind = col.index(product)
    except ValueError:
        while product not in col:
            if product[-1] == 's':
                t.setText(product + ' are not in the Excel sheet.\nType in a different product.')
            else:
                t.setText(product + ' is not in the Excel sheet.\nType in a different product.')
            m = p_win.getMouse()
            while not inside([m], _close_.getP1(), _graph_.getP2()):
                m = p_win.getMouse()
            if inside([m], _close_.getP1(), _close_.getP2()):
                p_win.close()
                return
            product = inp.getText().title().strip()
        ind = col.index(product)
    p_win.close()
    g_win = g.GraphWin("Prediction Graph", 800, 600)
    title = g.Text(g.Point(400, 15), 'Predictive Sales Graph for ' + product)
    title.draw(g_win)
    err_text = g.Text(g.Point(400, 30), '')
    err_text.draw(g_win)
    inp = g.Entry(g.Point(50, 590), 10)
    row = list(data.loc[ind, :].values)[2:]
    row_2 = list(data.loc[ind + 12, :].values)[2:]
    row_3 = list(data.loc[ind + 24, :].values)[2:]
    prices = [data.loc[ind, 'Price'], data.loc[ind + 12, 'Price'],
              data.loc[ind + 24, 'Price']]
    regs = []
    for i in range(len(row)):
        n = 3
        sum_x = sum(prices)
        sum_y = row[i] + row_2[i] + row_3[i]
        sum_xx = sum([x*x for x in prices])
        sum_xy = row[i] * prices[0] + row_2[i] * prices[1] + row_3[i] * prices[2]
        m = round((sum_xy - n * (sum_x / n) * (sum_y / n)) / (sum_xx - n * (sum_x / n) ** 2), 2)
        b = round(sum_y / n - m * (sum_x / n), 2)
        regs.append([m, b])
    inp.draw(g_win)
    inp.setText(str(prices[0]))
    _close_ = g.Rectangle(g.Point(0, 0), g.Point(50, 50))
    close_text = g.Text(g.Point(25, 25), 'Home')
    _close_.setFill('red')
    _close_.draw(g_win)
    _graph_ = g.Rectangle(g.Point(750, 550), g.Point(800, 600))
    _graph_.setFill("lightgreen")
    _graph_.draw(g_win)
    graph_text = g.Text(g.Point(775, 575), 'Graph')
    for item in [close_text, graph_text]:
        item.setFill('white')
        item.setStyle('bold')
        item.draw(g_win)
    make_graph(g_win, regs, float(inp.getText()), product)
    for item in [_close_, _graph_, graph_text, close_text, err_text, title, inp]:
        item.draw(g_win)
    m = g_win.getMouse()
    while True:
        while not inside([m], _close_.getP1(), _close_.getP2()) and not inside([m], _graph_.getP1(), _graph_.getP2()):
            m = g_win.getMouse()
        if inside([m], _close_.getP1(), _close_.getP2()):
            break
        else:
            try:
                num = float(inp.getText())
                err_text.setText('')
                make_graph(g_win, regs, num, product)
                for item in [_close_, _graph_, graph_text, close_text, err_text, title, inp]:
                    item.draw(g_win)
            except ValueError:
                err_text.setText('Make sure you type in a number.')
            m = g_win.getMouse()
    g_win.close()
    prediction()
    return


# clears the screen
def clear(win):
    for item in win.items[:]:
        item.undraw()
    win.update()


# makes the graph for the data provided
def make_graph(win, regs, price, product):
    clear(win)
    x_vals = [x - .5 for x in range(113, 719, 55)]
    y_axis = g.Line(g.Point(75, 50), g.Point(75, 525))
    y_axis.draw(win)
    x_axis = g.Line(g.Point(75, 525), g.Point(750, 525))
    x_axis.draw(win)
    a = 900
    for y in range(75, 526, 50):
        g.Line(g.Point(70, y), g.Point(75, y)).draw(win)
        g.Text(g.Point(50, y), str(a)).draw(win)
        a -= 100
    months = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec']
    sales = 0
    for i in range(12):
        x = x_vals[i]
        y = max(regs[i][0] * price + regs[i][1], 0)
        bar = g.Rectangle(g.Point(x - 27.5, 525 - y/2), g.Point(x + 27.5, 525))
        bar.draw(win)
        bar.setFill(g.color_rgb(0, 255, 0))
        if y >= 35:
            g.Text(bar.getCenter(), str(round(y))).draw(win)
        else:
            g.Text(g.Point(x, bar.getP1().getY() - 20), str(round(y))).draw(win)
        g.Text(g.Point(x, 540), months[i]).draw(win)
        sales += y
    prf = g.Text(g.Point(400, 560), "Money earned selling " + product.lower() + " at $" + str(price)
                 + ": $" + str(round(price * sales, 2)) + " annually")
    prf.draw(win)


# returns if any of the given points is inside a given box
def inside(points, tl, br):
    for point in points:
        x = point.getX()
        y = point.getY()
        if (br.getX() > x > tl.getX()) and (br.getY() > y > tl.getY()):
            return True
    return False


main()
