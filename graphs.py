import sqlite3
import pandas as pd
import matplotlib.pyplot as plt
import numpy as np

# Here is the answer to the first question
# It's the total ammount of tourists arrivals per year


def total(years):
    # Totals arrivals in 2011-2015
    values = []
    for year in years:
        # from the table of each year we select the cell that has the value of total arrivals in the whole year.
        table_name = "y%s" % year
        c.execute(
            'select "6" from %s where "1"="ΓΕΝΙΚΟ ΣΥΝΟΛΟ"' % table_name)
        db.commit()
        tr = c.fetchall()[0][0]
        num = round(float(tr))
        values.append(num)
    # wwe simply plot the graph
    plt.title('Total arrivals 2011-2015')
    plt.xlabel('Year')
    plt.ylabel('Tourists arrivals')
    plt.bar(years, values)


# In this function we check if a value is float
def is_float(arg):
    try:
        float(arg)
        return True
    except:
        return False


# This function is the answer to the second question of our project
def top_country(years):
    top_country = {}
    countries = []
    values = []
    for year in years:
        table_name = "y%s" % year
        c.execute('select "1","6" from %s' % table_name)
        db.commit()
        tr = c.fetchall()
        # we will loop through our data and store the name of the country and the total arrivals of tourists of this country where the arrivals are the most
        # compared to others
        _max = 0
        _country = None
        for t in tr:
            if is_float(t[1]) and (t[0] != "ΓΕΝΙΚΟ ΣΥΝΟΛΟ") and (t[0] != None) and (t[0] != "από τΙς οποίες:"):
                nm = float(t[1])
                number_of_tourists = int(round(nm))
                if number_of_tourists > _max:
                    _max = number_of_tourists
                    _country = t[0]
        top_country[year] = {_country: _max}
        countries.append(_country+"-"+str(year))
        values.append(_max)
    # We plot the graph
    plt.figure(figsize=(12, 7))
    plt.title('Total arrivals 2011-2015')
    plt.xlabel('Countries with largest tourists arrivals')
    plt.ylabel('Tourists arrivals')
    plt.bar(countries, values)


# turn the values of excel to integers.
def add(t):
    if is_float(t):
        return int(round(float(t)))
    else:
        return 0


# find the transports that were used every year by tourists in order to come to greece
def transports(years):
    transport = {}
    plane = 0
    car = 0
    train = 0
    boat = 0
    for year in years:
        # select the values that we need
        table_name = "y%s" % year
        c.execute('select "2","3","4","5" from %s' % table_name)
        db.commit()
        tr = c.fetchall()
        for t in tr:
            plane = plane + add(t[0])
            train = train + add(t[1])
            boat = boat + add(t[2])
            car = car + add(t[3])
            # dictionary with the year and the arrivals of each transport this year
        transport[year] = {"plane": plane,
                           "train": train, "ship": boat, "car": car}
    # Plot the graph using panda
    pd.DataFrame(transport).plot(kind='bar')
    plt.title('Means of transportation of tourists')
    plt.xlabel('Total in every year by mean')
    plt.ylabel('Tourists arrivals')


# this function help us plot the last graph
def bar_plot(ax, data, colors=None, total_width=0.8, single_width=1, legend=True):
    # Check if colors where provided, otherwhise use the default color cycle
    if colors is None:
        colors = plt.rcParams['axes.prop_cycle'].by_key()['color']
    # Number of bars per group
    n_bars = len(data)
    # The width of a single bar
    bar_width = total_width / n_bars
    # List containing handles for the drawn bars, used for the legend
    bars = []
    # Iterate over all data
    for i, (name, values) in enumerate(data.items()):
        # The offset in x direction of that bar
        x_offset = (i - n_bars / 2) * bar_width + bar_width / 2
        # Draw a bar for every value of that type
        for x, y in enumerate(values):
            bar = ax.bar(x + x_offset, y, width=bar_width *
                         single_width, color=colors[i % len(colors)])
        # Add a handle to the last drawn bar, which we'll need for the legend
        bars.append(bar[0])
    # Draw legend if we need
    if legend:
        ax.legend(bars, data.keys())


# Answer to the last question
# Arrivals per quarter per year
def quarters(years):
    c.execute('select "6" from "quarters" where "1"="ΓΕΝΙΚΟ ΣΥΝΟΛΟ"')
    db.commit()
    tr = c.fetchall()
    quarter = {}
    _quarter = 1
    i = 1
    total = 0
    # first we find the total ammount for each quarter
    for t in tr:
        total += int(round(t[0]))
        # we check every sheet in excel and every 3 sheets we have a quarter (there must be a better way to do that)
        if i % 3 == 0:
            quarter[_quarter] = total
            total = 0
            _quarter += 1
        i += 1
    quarters_per_year = {}
    # From here and all the way to the end, we try to plot the graph
    for i in range(1, 5):
        quarters_per_year['quarter %s' % i] = [quarter[i],
                                               quarter[i+4], quarter[i+8], quarter[i+12], quarter[i+16]]
    fig, ax = plt.subplots()
    bar_plot(ax, quarters_per_year, total_width=.8,
             single_width=.9)
    plt.title('Tourists arrivals per quarter per year')
    plt.xlabel('2011 - 2015')
    plt.ylabel('Arrivals')


# Connection to db
db = sqlite3.connect('data.db')
c = db.cursor()
years = [2011, 2012, 2013, 2014, 2015]

# With this command we check if the db is empty
c.execute('SELECT name FROM sqlite_master')
db.commit()
tr = c.fetchall()

# if tr is empty then we insert data to the db else we move on
if not tr:
    # Drop table in order to run multiple times
    c.execute("DROP TABLE IF EXISTS quarters")
    quarters_table = "quarters"

    # For every year
    # START DB CONFIGURATIONS
    for year in years:
        xl_name = "y%s.xls" % year
        xls = pd.ExcelFile(xl_name)
        names_sheet = xls.sheet_names
        name_sheet = names_sheet[-1]
        # dfs containes everything from the second table from the last sheet of xl file
        dfs = pd.read_excel(xl_name, sheet_name=name_sheet,
                            header=None, skiprows=range(1, 69))
        table_name = "y%s" % year
        # insert dfs to db
        dfs.to_sql(table_name, db, index=None, if_exists='replace')

        # for the last part of the exercise we need the total ammount of tourists from every quearter 2011-2015.
        # this will be stored in table "quarters"
        for name in names_sheet:
            df = pd.read_excel(xl_name, sheet_name=name,
                               header=None, skiprows=range(0, 63), nrows=6)
            df.to_sql(quarters_table, db, index=year, if_exists='append')
    # FINISHED DB CONFIGURATIONS

# We call every function
# This is the order of the questions in project
total(years)
top_country(years)
transports(years)
quarters(years)

#plt.bar(years, t_total)
plt.show()
